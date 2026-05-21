#[allow(dead_code)]
#[path = "../src/safety.rs"]
mod safety;

use anyhow::Result;
use safety::{scan_zip_bytes, OpcPackageKind};
use std::fs;
use std::io::{Cursor, Write};
use std::path::Path;
use tempfile::tempdir;
use zip::write::SimpleFileOptions;
use zip::ZipWriter;

#[test]
fn rejects_parent_traversal_paths() -> Result<()> {
    let temp = tempdir()?;

    let input_error = safety::resolve_input_path(temp.path(), "../secret.txt")
        .expect_err("parent traversal input should be rejected");
    assert!(input_error
        .to_string()
        .contains("parent directory traversal"));

    let output_error = safety::resolve_output_path(temp.path(), "out/../../secret.txt")
        .expect_err("parent traversal output should be rejected");
    assert!(output_error
        .to_string()
        .contains("parent directory traversal"));

    Ok(())
}

#[test]
fn rejects_symlink_output_when_platform_allows_symlink_creation() -> Result<()> {
    let temp = tempdir()?;
    let target = temp.path().join("target.txt");
    let link = temp.path().join("link.txt");
    fs::write(&target, b"target")?;

    if create_file_symlink(&target, &link).is_err() {
        return Ok(());
    }

    let resolve_error = safety::resolve_output_path(temp.path(), "link.txt")
        .expect_err("symlink output should be rejected");
    assert!(resolve_error.to_string().contains("symlink outputs"));

    let write_error = safety::atomic_write(&link, b"replacement")
        .expect_err("atomic write should reject symlink");
    assert!(write_error.to_string().contains("symlink outputs"));

    Ok(())
}

#[test]
fn scan_zip_flags_zip_slip_entry() -> Result<()> {
    let bytes = zip_bytes(&[("../evil.txt", b"bad".as_slice())])?;
    let report = scan_zip_bytes(&bytes, OpcPackageKind::Unknown)?;

    assert!(report.has_issue("ZIP_SLIP_ENTRY"));
    assert!(report.has_errors());
    Ok(())
}

#[test]
fn scan_zip_flags_duplicate_entry() -> Result<()> {
    let bytes = duplicate_empty_zip_bytes("docProps/core.xml");
    let report = scan_zip_bytes(&bytes, OpcPackageKind::Unknown)?;

    assert!(report.has_issue("ZIP_DUPLICATE_ENTRY"));
    assert!(report.has_errors());
    Ok(())
}

#[test]
fn scan_docx_flags_missing_required_parts() -> Result<()> {
    let bytes = zip_bytes(&[
        ("[Content_Types].xml", b"<Types/>".as_slice()),
        ("_rels/.rels", b"<Relationships/>".as_slice()),
    ])?;
    let report = scan_zip_bytes(&bytes, OpcPackageKind::Docx)?;

    assert!(report
        .issues
        .iter()
        .any(|issue| issue.code == "OPC_REQUIRED_PART_MISSING"
            && issue.part.as_deref() == Some("word/document.xml")));
    assert!(report.has_errors());
    Ok(())
}

#[test]
fn scan_zip_flags_external_relationships() -> Result<()> {
    let bytes = zip_bytes(&[
        ("[Content_Types].xml", b"<Types/>".as_slice()),
        (
            "_rels/.rels",
            br#"<Relationships><Relationship TargetMode="External" Target="https://example.test/template.dotm"/></Relationships>"#
                .as_slice(),
        ),
        ("word/document.xml", b"<w:document/>".as_slice()),
    ])?;
    let report = scan_zip_bytes(&bytes, OpcPackageKind::Docx)?;

    assert!(report.has_issue("ZIP_EXTERNAL_RELATIONSHIP"));
    assert!(report.has_errors());
    Ok(())
}

fn zip_bytes(entries: &[(&str, &[u8])]) -> Result<Vec<u8>> {
    let mut cursor = Cursor::new(Vec::new());
    {
        let mut writer = ZipWriter::new(&mut cursor);
        let options =
            SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated);
        for (name, bytes) in entries {
            writer.start_file(*name, options)?;
            writer.write_all(bytes)?;
        }
        writer.finish()?;
    }
    Ok(cursor.into_inner())
}

fn duplicate_empty_zip_bytes(name: &str) -> Vec<u8> {
    let mut bytes = Vec::new();
    let first_offset = bytes.len() as u32;
    write_local_file_header(&mut bytes, name);
    let second_offset = bytes.len() as u32;
    write_local_file_header(&mut bytes, name);

    let central_directory_offset = bytes.len() as u32;
    write_central_directory_header(&mut bytes, name, first_offset);
    write_central_directory_header(&mut bytes, name, second_offset);
    let central_directory_size = bytes.len() as u32 - central_directory_offset;

    write_u32(&mut bytes, 0x0605_4b50);
    write_u16(&mut bytes, 0);
    write_u16(&mut bytes, 0);
    write_u16(&mut bytes, 2);
    write_u16(&mut bytes, 2);
    write_u32(&mut bytes, central_directory_size);
    write_u32(&mut bytes, central_directory_offset);
    write_u16(&mut bytes, 0);

    bytes
}

fn write_local_file_header(bytes: &mut Vec<u8>, name: &str) {
    write_u32(bytes, 0x0403_4b50);
    write_u16(bytes, 20);
    write_u16(bytes, 0);
    write_u16(bytes, 0);
    write_u16(bytes, 0);
    write_u16(bytes, 0);
    write_u32(bytes, 0);
    write_u32(bytes, 0);
    write_u32(bytes, 0);
    write_u16(bytes, name.len() as u16);
    write_u16(bytes, 0);
    bytes.extend_from_slice(name.as_bytes());
}

fn write_central_directory_header(bytes: &mut Vec<u8>, name: &str, local_offset: u32) {
    write_u32(bytes, 0x0201_4b50);
    write_u16(bytes, 20);
    write_u16(bytes, 20);
    write_u16(bytes, 0);
    write_u16(bytes, 0);
    write_u16(bytes, 0);
    write_u16(bytes, 0);
    write_u32(bytes, 0);
    write_u32(bytes, 0);
    write_u32(bytes, 0);
    write_u16(bytes, name.len() as u16);
    write_u16(bytes, 0);
    write_u16(bytes, 0);
    write_u16(bytes, 0);
    write_u16(bytes, 0);
    write_u32(bytes, 0);
    write_u32(bytes, local_offset);
    bytes.extend_from_slice(name.as_bytes());
}

fn write_u16(bytes: &mut Vec<u8>, value: u16) {
    bytes.extend_from_slice(&value.to_le_bytes());
}

fn write_u32(bytes: &mut Vec<u8>, value: u32) {
    bytes.extend_from_slice(&value.to_le_bytes());
}

#[cfg(windows)]
fn create_file_symlink(target: &Path, link: &Path) -> std::io::Result<()> {
    std::os::windows::fs::symlink_file(target, link)
}

#[cfg(unix)]
fn create_file_symlink(target: &Path, link: &Path) -> std::io::Result<()> {
    std::os::unix::fs::symlink(target, link)
}
