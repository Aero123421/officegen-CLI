use anyhow::{bail, Context, Result};
use std::collections::BTreeSet;
use std::fs::{self, File};
use std::io::{Cursor, Read, Seek, Write};
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicU64, Ordering};
use std::time::{SystemTime, UNIX_EPOCH};
use zip::ZipArchive;

static TEMP_COUNTER: AtomicU64 = AtomicU64::new(0);

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum OpcPackageKind {
    Unknown,
    Docx,
    Pptx,
    Xlsx,
}

impl OpcPackageKind {
    pub fn from_path(path: &Path) -> Self {
        match path
            .extension()
            .and_then(|extension| extension.to_str())
            .unwrap_or_default()
            .to_ascii_lowercase()
            .as_str()
        {
            "docx" => Self::Docx,
            "pptx" => Self::Pptx,
            "xlsx" => Self::Xlsx,
            _ => Self::Unknown,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SafetySeverity {
    Info,
    Warning,
    Error,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ZipSafetyIssue {
    pub code: &'static str,
    pub severity: SafetySeverity,
    pub part: Option<String>,
    pub message: String,
}

impl ZipSafetyIssue {
    fn error(
        code: &'static str,
        part: impl Into<Option<String>>,
        message: impl Into<String>,
    ) -> Self {
        Self {
            code,
            severity: SafetySeverity::Error,
            part: part.into(),
            message: message.into(),
        }
    }

    fn warning(
        code: &'static str,
        part: impl Into<Option<String>>,
        message: impl Into<String>,
    ) -> Self {
        Self {
            code,
            severity: SafetySeverity::Warning,
            part: part.into(),
            message: message.into(),
        }
    }

    fn info(
        code: &'static str,
        part: impl Into<Option<String>>,
        message: impl Into<String>,
    ) -> Self {
        Self {
            code,
            severity: SafetySeverity::Info,
            part: part.into(),
            message: message.into(),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ZipSafetyReport {
    pub entries: BTreeSet<String>,
    pub issues: Vec<ZipSafetyIssue>,
}

impl ZipSafetyReport {
    #[allow(dead_code)]
    pub fn has_errors(&self) -> bool {
        self.issues
            .iter()
            .any(|issue| issue.severity == SafetySeverity::Error)
    }

    #[allow(dead_code)]
    pub fn has_issue(&self, code: &str) -> bool {
        self.issues.iter().any(|issue| issue.code == code)
    }
}

pub fn resolve_input_path(root: &Path, value: &str) -> Result<PathBuf> {
    let root = canonical_root(root)?;
    let candidate = resolve_root_relative(&root, value)?;
    if !candidate.exists() {
        bail!("INPUT_NOT_FOUND: {}", redacted_path(&candidate));
    }

    let canonical = candidate
        .canonicalize()
        .with_context(|| format!("failed to canonicalize input {}", redacted_path(&candidate)))?;
    if !canonical.starts_with(&root) {
        bail!("SECURITY_PATH_OUTSIDE_ROOT: input resolves outside project");
    }
    Ok(canonical)
}

pub fn resolve_output_path(root: &Path, value: &str) -> Result<PathBuf> {
    let root = canonical_root(root)?;
    let candidate = resolve_root_relative(&root, value)?;

    deny_existing_symlink_components(&root, Path::new(value))?;

    let mut probe = candidate
        .parent()
        .map(Path::to_path_buf)
        .unwrap_or_else(|| root.clone());
    while !probe.exists() {
        if !probe.pop() {
            break;
        }
    }

    if probe.exists() {
        let canonical_parent = probe.canonicalize().with_context(|| {
            format!(
                "failed to canonicalize output parent {}",
                redacted_path(&probe)
            )
        })?;
        if !canonical_parent.starts_with(&root) {
            bail!("SECURITY_PATH_OUTSIDE_ROOT: output parent resolves outside project");
        }
    }

    if candidate.exists() && is_symlink(&candidate)? {
        bail!("SECURITY_PATH_OUTSIDE_ROOT: symlink outputs are not allowed");
    }

    Ok(candidate)
}

pub fn atomic_write(path: &Path, data: &[u8]) -> Result<()> {
    if path.exists() && is_symlink(path)? {
        bail!("SECURITY_PATH_OUTSIDE_ROOT: symlink outputs are not allowed");
    }

    let parent = path.parent().unwrap_or_else(|| Path::new("."));
    fs::create_dir_all(parent).with_context(|| {
        format!(
            "failed to create output directory {}",
            redacted_path(parent)
        )
    })?;

    let temp = temporary_sibling_path(path);
    let write_result = (|| -> Result<()> {
        let mut file = File::create(&temp)
            .with_context(|| format!("failed to create temp file {}", redacted_path(&temp)))?;
        file.write_all(data)
            .with_context(|| format!("failed to write temp file {}", redacted_path(&temp)))?;
        file.sync_all()
            .with_context(|| format!("failed to sync temp file {}", redacted_path(&temp)))?;
        fs::rename(&temp, path).with_context(|| {
            format!(
                "failed to move temp file {} into {}",
                redacted_path(&temp),
                redacted_path(path)
            )
        })?;
        Ok(())
    })();

    if write_result.is_err() {
        let _ = fs::remove_file(&temp);
    }

    write_result
}

pub fn scan_zip_file(path: &Path) -> Result<ZipSafetyReport> {
    let bytes =
        fs::read(path).with_context(|| format!("failed to read {}", redacted_path(path)))?;
    scan_zip_bytes(&bytes, OpcPackageKind::from_path(path))
}

pub fn scan_zip_bytes(bytes: &[u8], kind: OpcPackageKind) -> Result<ZipSafetyReport> {
    let duplicate_names = central_directory_duplicate_names(bytes);
    let mut zip = ZipArchive::new(Cursor::new(bytes))?;
    let mut entries = BTreeSet::new();
    let mut issues = duplicate_names
        .iter()
        .map(|name| {
            ZipSafetyIssue::error(
                "ZIP_DUPLICATE_ENTRY",
                Some(name.clone()),
                "duplicate ZIP entry name",
            )
        })
        .collect::<Vec<_>>();

    for index in 0..zip.len() {
        let mut file = zip.by_index(index)?;
        let name = file.name().to_string();
        scan_entry_name(&name, &mut issues);

        if !entries.insert(name.clone()) && !duplicate_names.contains(&name) {
            issues.push(ZipSafetyIssue::error(
                "ZIP_DUPLICATE_ENTRY",
                Some(name.clone()),
                "duplicate ZIP entry name",
            ));
        }

        scan_sensitive_part(&name, &mut issues);

        if is_relationship_part(&name) {
            let mut bytes = Vec::new();
            file.read_to_end(&mut bytes)?;
            let xml = String::from_utf8_lossy(&bytes);
            scan_relationship_part(&name, &xml, &mut issues);
        }
    }

    scan_required_opc_parts(kind, &entries, &mut issues);

    Ok(ZipSafetyReport { entries, issues })
}

#[allow(dead_code)]
pub fn scan_zip_reader<R: Read + Seek>(reader: R, kind: OpcPackageKind) -> Result<ZipSafetyReport> {
    let mut reader = reader;
    let mut bytes = Vec::new();
    reader.read_to_end(&mut bytes)?;
    scan_zip_bytes(&bytes, kind)
}

fn canonical_root(root: &Path) -> Result<PathBuf> {
    root.canonicalize()
        .with_context(|| format!("failed to canonicalize root {}", redacted_path(root)))
}

fn resolve_root_relative(root: &Path, value: &str) -> Result<PathBuf> {
    validate_user_path(value)?;

    let mut candidate = root.to_path_buf();
    for segment in split_user_path(value) {
        if segment == "." || segment.is_empty() {
            continue;
        }
        candidate.push(segment);
    }
    Ok(candidate)
}

fn validate_user_path(value: &str) -> Result<()> {
    if value.trim().is_empty() {
        bail!("SECURITY_PATH_OUTSIDE_ROOT: empty paths are not allowed");
    }
    if Path::new(value).is_absolute()
        || value.starts_with('/')
        || value.starts_with('\\')
        || value.starts_with("//")
        || value.starts_with("\\\\")
    {
        bail!("SECURITY_PATH_OUTSIDE_ROOT: absolute paths are not allowed");
    }
    if has_windows_drive_prefix(value) {
        bail!("SECURITY_PATH_OUTSIDE_ROOT: Windows drive paths are not allowed");
    }
    if split_user_path(value).any(|segment| segment == "..") {
        bail!("SECURITY_PATH_OUTSIDE_ROOT: parent directory traversal is not allowed");
    }
    Ok(())
}

fn split_user_path(value: &str) -> impl Iterator<Item = &str> {
    value.split(['/', '\\'])
}

fn deny_existing_symlink_components(root: &Path, value: &Path) -> Result<()> {
    let mut current = root.to_path_buf();
    for component in value.to_string_lossy().split(['/', '\\']) {
        if component.is_empty() || component == "." {
            continue;
        }
        current.push(component);
        if current.exists() && is_symlink(&current)? {
            bail!("SECURITY_PATH_OUTSIDE_ROOT: symlink outputs are not allowed");
        }
    }
    Ok(())
}

fn is_symlink(path: &Path) -> Result<bool> {
    Ok(fs::symlink_metadata(path)?.file_type().is_symlink())
}

fn temporary_sibling_path(path: &Path) -> PathBuf {
    let counter = TEMP_COUNTER.fetch_add(1, Ordering::Relaxed);
    let nanos = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|duration| duration.as_nanos())
        .unwrap_or_default();
    let file_name = path
        .file_name()
        .and_then(|value| value.to_str())
        .unwrap_or("officegen-output");
    path.with_file_name(format!(
        ".{file_name}.officegen-tmp-{}-{nanos}-{counter}",
        std::process::id()
    ))
}

fn central_directory_duplicate_names(bytes: &[u8]) -> BTreeSet<String> {
    let mut seen = BTreeSet::new();
    let mut duplicates = BTreeSet::new();

    let Some(eocd_offset) = find_eocd(bytes) else {
        return duplicates;
    };
    if eocd_offset + 22 > bytes.len() {
        return duplicates;
    }

    let entry_count = read_u16_le(bytes, eocd_offset + 10).unwrap_or_default() as usize;
    let central_directory_size = read_u32_le(bytes, eocd_offset + 12).unwrap_or_default() as usize;
    let central_directory_offset =
        read_u32_le(bytes, eocd_offset + 16).unwrap_or_default() as usize;
    let central_directory_end = central_directory_offset.saturating_add(central_directory_size);
    if central_directory_offset > bytes.len() || central_directory_end > bytes.len() {
        return duplicates;
    }

    let mut offset = central_directory_offset;
    for _ in 0..entry_count {
        if offset + 46 > central_directory_end || read_u32_le(bytes, offset) != Some(0x0201_4b50) {
            break;
        }

        let name_len = read_u16_le(bytes, offset + 28).unwrap_or_default() as usize;
        let extra_len = read_u16_le(bytes, offset + 30).unwrap_or_default() as usize;
        let comment_len = read_u16_le(bytes, offset + 32).unwrap_or_default() as usize;
        let name_start = offset + 46;
        let name_end = name_start.saturating_add(name_len);
        let next_offset = name_end
            .saturating_add(extra_len)
            .saturating_add(comment_len);

        if name_end > central_directory_end || next_offset > central_directory_end {
            break;
        }

        let name = String::from_utf8_lossy(&bytes[name_start..name_end]).to_string();
        if !seen.insert(name.clone()) {
            duplicates.insert(name);
        }

        offset = next_offset;
    }

    duplicates
}

fn find_eocd(bytes: &[u8]) -> Option<usize> {
    let search_start = bytes.len().saturating_sub(65_557);
    (search_start..bytes.len().saturating_sub(3))
        .rev()
        .find(|offset| bytes[*offset..].starts_with(&[0x50, 0x4b, 0x05, 0x06]))
}

fn read_u16_le(bytes: &[u8], offset: usize) -> Option<u16> {
    let slice = bytes.get(offset..offset + 2)?;
    Some(u16::from_le_bytes([slice[0], slice[1]]))
}

fn read_u32_le(bytes: &[u8], offset: usize) -> Option<u32> {
    let slice = bytes.get(offset..offset + 4)?;
    Some(u32::from_le_bytes([slice[0], slice[1], slice[2], slice[3]]))
}

fn scan_entry_name(name: &str, issues: &mut Vec<ZipSafetyIssue>) {
    let normalized = name.replace('\\', "/");

    if normalized.is_empty() {
        issues.push(ZipSafetyIssue::error(
            "ZIP_EMPTY_ENTRY",
            Some(name.to_string()),
            "empty ZIP entry name",
        ));
        return;
    }

    if normalized.starts_with('/') || normalized.starts_with("//") {
        issues.push(ZipSafetyIssue::error(
            "ZIP_ABSOLUTE_ENTRY",
            Some(name.to_string()),
            "absolute ZIP entry path",
        ));
    }

    if has_windows_drive_prefix(&normalized) {
        issues.push(ZipSafetyIssue::error(
            "ZIP_WINDOWS_DRIVE_ENTRY",
            Some(name.to_string()),
            "Windows drive-qualified ZIP entry path",
        ));
    }

    if normalized.split('/').any(|segment| segment == "..") {
        issues.push(ZipSafetyIssue::error(
            "ZIP_SLIP_ENTRY",
            Some(name.to_string()),
            "ZIP entry contains parent traversal",
        ));
    }
}

fn scan_sensitive_part(name: &str, issues: &mut Vec<ZipSafetyIssue>) {
    let normalized = name.replace('\\', "/");
    let lower = normalized.to_ascii_lowercase();

    if lower.contains("vbaproject.bin") {
        issues.push(ZipSafetyIssue::error(
            "ZIP_MACRO_PART",
            Some(name.to_string()),
            "macro project part is not allowed",
        ));
    }

    if lower.contains("/embeddings/") {
        if lower.ends_with(".xlsx") {
            issues.push(ZipSafetyIssue::info(
                "ZIP_EMBEDDED_WORKBOOK",
                Some(name.to_string()),
                "embedded workbook part",
            ));
        } else {
            issues.push(ZipSafetyIssue::warning(
                "ZIP_EMBEDDED_OBJECT",
                Some(name.to_string()),
                "embedded object part",
            ));
        }
    }
}

fn scan_relationship_part(name: &str, xml: &str, issues: &mut Vec<ZipSafetyIssue>) {
    let compact = xml
        .chars()
        .filter(|character| !character.is_whitespace())
        .collect::<String>()
        .to_ascii_lowercase();

    if compact.contains("targetmode=\"external\"")
        || compact.contains("targetmode='external'")
        || compact.contains("target=\"http://")
        || compact.contains("target='http://")
        || compact.contains("target=\"https://")
        || compact.contains("target='https://")
        || compact.contains("target=\"file:")
        || compact.contains("target='file:")
        || compact.contains("target=\"//")
        || compact.contains("target='//")
    {
        issues.push(ZipSafetyIssue::error(
            "ZIP_EXTERNAL_RELATIONSHIP",
            Some(name.to_string()),
            "external OPC relationship target is not allowed",
        ));
    }
}

fn scan_required_opc_parts(
    kind: OpcPackageKind,
    entries: &BTreeSet<String>,
    issues: &mut Vec<ZipSafetyIssue>,
) {
    require_exact("[Content_Types].xml", entries, issues);
    require_exact("_rels/.rels", entries, issues);

    match kind {
        OpcPackageKind::Docx => {
            require_exact("word/document.xml", entries, issues);
        }
        OpcPackageKind::Pptx => {
            require_exact("ppt/presentation.xml", entries, issues);
            require_exact("ppt/_rels/presentation.xml.rels", entries, issues);
            require_prefix_suffix(
                "ppt/slides/slide*.xml",
                "ppt/slides/slide",
                ".xml",
                entries,
                issues,
            );
        }
        OpcPackageKind::Xlsx => {
            require_exact("xl/workbook.xml", entries, issues);
            require_exact("xl/_rels/workbook.xml.rels", entries, issues);
            require_prefix_suffix(
                "xl/worksheets/sheet*.xml",
                "xl/worksheets/sheet",
                ".xml",
                entries,
                issues,
            );
        }
        OpcPackageKind::Unknown => {}
    }
}

fn require_exact(part: &'static str, entries: &BTreeSet<String>, issues: &mut Vec<ZipSafetyIssue>) {
    if !entries.contains(part) {
        issues.push(ZipSafetyIssue::error(
            "OPC_REQUIRED_PART_MISSING",
            Some(part.to_string()),
            "required OPC part is missing",
        ));
    }
}

fn require_prefix_suffix(
    label: &'static str,
    prefix: &str,
    suffix: &str,
    entries: &BTreeSet<String>,
    issues: &mut Vec<ZipSafetyIssue>,
) {
    if !entries
        .iter()
        .any(|entry| entry.starts_with(prefix) && entry.ends_with(suffix))
    {
        issues.push(ZipSafetyIssue::error(
            "OPC_REQUIRED_PART_MISSING",
            Some(label.to_string()),
            "required OPC part is missing",
        ));
    }
}

fn is_relationship_part(name: &str) -> bool {
    let normalized = name.replace('\\', "/").to_ascii_lowercase();
    normalized.ends_with(".rels") && normalized.contains("_rels/")
}

fn has_windows_drive_prefix(value: &str) -> bool {
    let bytes = value.as_bytes();
    bytes.len() >= 2 && bytes[1] == b':' && bytes[0].is_ascii_alphabetic()
}

fn redacted_path(path: &Path) -> String {
    path.file_name()
        .and_then(|value| value.to_str())
        .unwrap_or("<path>")
        .to_string()
}
