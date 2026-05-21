#![allow(dead_code)]

use anyhow::{anyhow, bail, Result};
use regex::Regex;
use serde_json::{json, Value};
use sha2::{Digest, Sha256};
use std::fs::{self, File};
use std::io::{Cursor, Read, Write};
use std::path::Path;
use zip::write::SimpleFileOptions;
use zip::{ZipArchive, ZipWriter};

#[derive(Clone, Debug)]
struct SlideSpec {
    title: String,
    blocks: Vec<Value>,
}

#[derive(Clone, Debug)]
struct OoxmlSummary {
    format: String,
    parts: Vec<String>,
    texts: Vec<String>,
    tables: Vec<Vec<Vec<String>>>,
    images: usize,
    charts: usize,
    object_map: Vec<Value>,
}

pub fn write_ir_pptx(path: &Path, ir: &Value, fallback_title: &str) -> Result<()> {
    let slides = slide_specs(ir, fallback_title);
    let mut buffer = Cursor::new(Vec::new());
    let mut writer = ZipWriter::new(&mut buffer);
    let options = SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated);

    zip_file(
        &mut writer,
        "[Content_Types].xml",
        pptx_content_types(slides.len()),
        options,
    )?;
    zip_file(
        &mut writer,
        "_rels/.rels",
        rels("officeDocument", "ppt/presentation.xml"),
        options,
    )?;
    zip_file(
        &mut writer,
        "ppt/presentation.xml",
        pptx_presentation_xml(slides.len()),
        options,
    )?;
    zip_file(
        &mut writer,
        "ppt/_rels/presentation.xml.rels",
        pptx_presentation_rels(slides.len()),
        options,
    )?;
    for (index, slide) in slides.iter().enumerate() {
        zip_file(
            &mut writer,
            &format!("ppt/slides/slide{}.xml", index + 1),
            pptx_slide_xml(index + 1, slide),
            options,
        )?;
    }
    writer.finish()?;
    crate::safety::atomic_write(path, &buffer.into_inner())?;
    Ok(())
}

pub fn write_ir_docx(path: &Path, ir: &Value, fallback_title: &str) -> Result<()> {
    let mut buffer = Cursor::new(Vec::new());
    let mut writer = ZipWriter::new(&mut buffer);
    let options = SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated);

    zip_file(
        &mut writer,
        "[Content_Types].xml",
        docx_content_types(),
        options,
    )?;
    zip_file(
        &mut writer,
        "_rels/.rels",
        rels("officeDocument", "word/document.xml"),
        options,
    )?;
    zip_file(
        &mut writer,
        "word/document.xml",
        docx_document_xml(ir, fallback_title),
        options,
    )?;
    writer.finish()?;
    crate::safety::atomic_write(path, &buffer.into_inner())?;
    Ok(())
}

pub fn inspect_ooxml(path: &Path, format: &str) -> Result<Value> {
    crate::safety::scan_zip_file(path).and_then(|report| {
        if let Some(issue) = report
            .issues
            .iter()
            .find(|issue| issue.severity == crate::safety::SafetySeverity::Error)
        {
            bail!(
                "SECURITY_ZIP_UNSAFE: {}{}",
                issue.code,
                issue
                    .part
                    .as_ref()
                    .map(|part| format!(" in {part}"))
                    .unwrap_or_default()
            );
        }
        Ok(())
    })?;

    let summary = summarize_ooxml(path, format)?;
    let text_preview = summary
        .texts
        .join(" ")
        .chars()
        .take(2000)
        .collect::<String>();
    Ok(json!({
        "schema": "officegen.inspect.result@1.2",
        "format": format,
        "trusted": {
            "summary": {
                "format": format,
                "parts": summary.parts.len(),
                "textObjects": summary.texts.iter().filter(|text| !text.trim().is_empty()).count(),
                "tables": summary.tables.len(),
                "images": summary.images,
                "charts": summary.charts,
                "characters": summary.texts.iter().map(|s| s.chars().count()).sum::<usize>(),
                "sha256": sha256_file(path)?,
                "semanticPresence": presence_json(&summary)
            }
        },
        "untrusted": {"textPreview": text_preview, "parts": summary.parts},
        "objectMap": summary.object_map,
        "package": {"parts": summary.parts}
    }))
}

pub fn verify_hints(path: &Path, format: &str) -> Result<(Value, Vec<Value>)> {
    let summary = summarize_ooxml(path, format)?;
    let mut issues = Vec::new();
    if summary.texts.iter().all(|text| text.trim().is_empty())
        && summary.tables.is_empty()
        && summary.images == 0
        && summary.charts == 0
    {
        issues.push(json!({
            "code": "OOXML_EMPTY_CONTENT_HINT",
            "severity": "warning",
            "message": "No text, table, image, or chart-like content was detected by portable inspection."
        }));
    }
    for text in &summary.texts {
        let limit = if format == "pptx" { 260 } else { 1400 };
        if text.chars().count() > limit {
            issues.push(json!({
                "code": "TEXT_OVERFLOW_HINT",
                "severity": "warning",
                "characters": text.chars().count(),
                "message": "A text run is long enough to deserve native layout verification for overflow."
            }));
            break;
        }
    }
    if summary
        .tables
        .iter()
        .flatten()
        .flatten()
        .any(|cell| cell.chars().count() > 140)
    {
        issues.push(json!({
            "code": "TABLE_CELL_OVERFLOW_HINT",
            "severity": "warning",
            "message": "A table cell contains long text; verify layout in the target Office renderer."
        }));
    }
    Ok((presence_json(&summary), issues))
}

pub fn semantic_diff(before: &Path, after: &Path, format: &str) -> Result<Value> {
    let before_summary = summarize_ooxml(before, format)?;
    let after_summary = summarize_ooxml(after, format)?;
    let before_tables = table_fingerprint(&before_summary.tables);
    let after_tables = table_fingerprint(&after_summary.tables);
    let before_text = before_summary.texts.join("\n");
    let after_text = after_summary.texts.join("\n");

    Ok(json!({
        "changedTextObjects": if before_text == after_text { 0 } else { 1 },
        "textChanged": before_text != after_text,
        "tableChanged": before_tables != after_tables,
        "imagePresenceChanged": before_summary.images != after_summary.images,
        "chartPresenceChanged": before_summary.charts != after_summary.charts,
        "presenceBefore": presence_json(&before_summary),
        "presenceAfter": presence_json(&after_summary),
        "tablesBefore": before_summary.tables.len(),
        "tablesAfter": after_summary.tables.len(),
        "imagesBefore": before_summary.images,
        "imagesAfter": after_summary.images,
        "chartsBefore": before_summary.charts,
        "chartsAfter": after_summary.charts
    }))
}

pub fn view_html(inspected: &Value) -> String {
    let summary = inspected
        .pointer("/trusted/summary")
        .cloned()
        .unwrap_or_else(|| json!({}));
    let object_map = inspected
        .get("objectMap")
        .and_then(Value::as_array)
        .cloned()
        .unwrap_or_default();
    let mut objects = String::new();
    for object in object_map.iter().take(80) {
        let object_type = object
            .get("type")
            .and_then(Value::as_str)
            .unwrap_or("object");
        let source = object
            .get("sourcePath")
            .and_then(Value::as_str)
            .unwrap_or("");
        let preview = object
            .get("textPreview")
            .and_then(Value::as_str)
            .unwrap_or("");
        objects.push_str(&format!(
            "<section class=\"object\"><div><strong>{}</strong><span>{}</span></div><p>{}</p>{}</section>",
            html_escape(object_type),
            html_escape(source),
            html_escape(preview),
            table_html(object)
        ));
    }
    format!(
        "<!doctype html><meta charset=\"utf-8\"><title>Officegen OOXML View</title><style>body{{font:14px/1.45 system-ui,Segoe UI,sans-serif;margin:0;background:#f7f7f5;color:#202124}}header{{position:sticky;top:0;background:white;border-bottom:1px solid #ddd;padding:16px 24px}}main{{padding:20px 24px;max-width:1120px}}.chips{{display:flex;gap:8px;flex-wrap:wrap;margin-top:10px}}.chip{{border:1px solid #ccc;background:#fff;padding:4px 8px;border-radius:6px}}.object{{background:white;border:1px solid #ddd;border-radius:8px;padding:12px 14px;margin:0 0 12px}}.object div{{display:flex;gap:12px;justify-content:space-between;color:#5f6368}}.object p{{white-space:pre-wrap;margin:8px 0 0}}table{{border-collapse:collapse;margin-top:10px}}td{{border:1px solid #d0d0d0;padding:5px 8px}}</style><header><h1>OOXML Preview</h1><div class=\"chips\">{}</div></header><main>{}</main>",
        summary_chips(&summary),
        objects
    )
}

pub fn view_svg(inspected: &Value, width: usize, height: usize) -> String {
    let summary = inspected
        .pointer("/trusted/summary")
        .cloned()
        .unwrap_or_else(|| json!({}));
    let mut body = String::new();
    body.push_str(&format!(
        "<text x=\"32\" y=\"44\" font-size=\"24\" font-family=\"Arial, sans-serif\">{} preview</text>",
        html_escape(summary.get("format").and_then(Value::as_str).unwrap_or("OOXML"))
    ));
    let chips = [
        (
            "Text",
            summary
                .get("textObjects")
                .and_then(Value::as_u64)
                .unwrap_or(0),
        ),
        (
            "Tables",
            summary.get("tables").and_then(Value::as_u64).unwrap_or(0),
        ),
        (
            "Images",
            summary.get("images").and_then(Value::as_u64).unwrap_or(0),
        ),
        (
            "Charts",
            summary.get("charts").and_then(Value::as_u64).unwrap_or(0),
        ),
    ];
    for (i, (label, count)) in chips.iter().enumerate() {
        let x = 32 + i * 132;
        body.push_str(&format!("<rect x=\"{x}\" y=\"64\" width=\"112\" height=\"34\" rx=\"5\" fill=\"#eef2f4\" stroke=\"#b9c2c8\"/><text x=\"{}\" y=\"86\" font-size=\"14\" font-family=\"Arial, sans-serif\">{}: {}</text>", x + 12, html_escape(label), count));
    }
    let object_map = inspected
        .get("objectMap")
        .and_then(Value::as_array)
        .cloned()
        .unwrap_or_default();
    for (i, object) in object_map.iter().take(12).enumerate() {
        let y = 132 + i * 30;
        let object_type = object
            .get("type")
            .and_then(Value::as_str)
            .unwrap_or("object");
        let preview = object
            .get("textPreview")
            .and_then(Value::as_str)
            .unwrap_or("");
        body.push_str(&format!(
            "<text x=\"32\" y=\"{y}\" font-size=\"16\" font-family=\"Arial, sans-serif\">{}: {}</text>",
            html_escape(object_type),
            html_escape(&preview.chars().take(86).collect::<String>())
        ));
    }
    format!("<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"{width}\" height=\"{height}\"><rect width=\"100%\" height=\"100%\" fill=\"white\"/>{body}</svg>")
}

pub fn apply_table_cell_xml(
    format: &str,
    part: &str,
    xml: &str,
    op: &Value,
) -> Result<Option<String>> {
    if format == "docx" && part != "word/document.xml" {
        return Ok(None);
    }
    if format == "pptx" && !part.starts_with("ppt/slides/") {
        return Ok(None);
    }
    let replacement = op
        .get("text")
        .and_then(Value::as_str)
        .or_else(|| op.get("value").and_then(Value::as_str))
        .unwrap_or("");
    let table_index = op
        .pointer("/selector/tableIndex")
        .and_then(Value::as_u64)
        .or_else(|| op.get("tableIndex").and_then(Value::as_u64))
        .unwrap_or(0) as usize;
    let row_index = index_arg(op, "rowIndex", "row")?;
    let col_index =
        index_arg(op, "columnIndex", "col").or_else(|_| index_arg(op, "columnIndex", "column"))?;
    match format {
        "docx" => replace_table_cell(
            xml,
            table_index,
            row_index,
            col_index,
            replacement,
            "w:tbl",
            "w:tr",
            "w:tc",
            "w:t",
        ),
        "pptx" => replace_table_cell(
            xml,
            table_index,
            row_index,
            col_index,
            replacement,
            "a:tbl",
            "a:tr",
            "a:tc",
            "a:t",
        ),
        _ => Ok(None),
    }
}

fn slide_specs(ir: &Value, fallback_title: &str) -> Vec<SlideSpec> {
    if let Some(slides) = ir.get("slides").and_then(Value::as_array) {
        let specs = slides
            .iter()
            .enumerate()
            .map(|(index, slide)| SlideSpec {
                title: slide
                    .get("title")
                    .and_then(Value::as_str)
                    .map(str::to_string)
                    .unwrap_or_else(|| format!("Slide {}", index + 1)),
                blocks: slide
                    .get("blocks")
                    .and_then(Value::as_array)
                    .cloned()
                    .unwrap_or_default(),
            })
            .collect::<Vec<_>>();
        if !specs.is_empty() {
            return specs;
        }
    }
    let mut specs = ir
        .get("sections")
        .and_then(Value::as_array)
        .unwrap_or(&Vec::new())
        .iter()
        .enumerate()
        .map(|(index, section)| SlideSpec {
            title: section
                .get("title")
                .and_then(Value::as_str)
                .map(str::to_string)
                .unwrap_or_else(|| {
                    if index == 0 {
                        fallback_title.to_string()
                    } else {
                        format!("Section {}", index + 1)
                    }
                }),
            blocks: section
                .get("blocks")
                .and_then(Value::as_array)
                .cloned()
                .unwrap_or_default(),
        })
        .collect::<Vec<_>>();
    if specs.is_empty() {
        specs.push(SlideSpec {
            title: fallback_title.to_string(),
            blocks: vec![json!({"type": "paragraph", "text": collect_ir_text(ir)})],
        });
    }
    specs
}

fn docx_document_xml(ir: &Value, fallback_title: &str) -> String {
    let mut body = String::new();
    body.push_str(&docx_paragraph(fallback_title, "Title"));
    if let Some(blocks) = ir.get("blocks").and_then(Value::as_array) {
        for block in blocks {
            if block_text(block).trim() == fallback_title.trim()
                && matches!(
                    block.get("type").and_then(Value::as_str),
                    Some("heading" | "title")
                )
            {
                continue;
            }
            body.push_str(&docx_block_xml(block));
        }
    }
    for (section_index, section) in ir
        .get("sections")
        .and_then(Value::as_array)
        .unwrap_or(&Vec::new())
        .iter()
        .enumerate()
    {
        if let Some(title) = section.get("title").and_then(Value::as_str) {
            if !(section_index == 0 && title.trim() == fallback_title.trim()) {
                body.push_str(&docx_paragraph(title, "Heading1"));
            }
        }
        for block in section
            .get("blocks")
            .and_then(Value::as_array)
            .unwrap_or(&Vec::new())
        {
            if block_text(block).trim() == fallback_title.trim()
                && matches!(
                    block.get("type").and_then(Value::as_str),
                    Some("heading" | "title")
                )
            {
                continue;
            }
            body.push_str(&docx_block_xml(block));
        }
    }
    if body.trim().is_empty() {
        body.push_str(&docx_paragraph(&collect_ir_text(ir), "Normal"));
    }
    format!("<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>{body}<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\"/></w:sectPr></w:body></w:document>")
}

fn docx_block_xml(block: &Value) -> String {
    let kind = block
        .get("type")
        .and_then(Value::as_str)
        .unwrap_or("paragraph");
    match kind {
        "heading" | "title" => docx_paragraph(&block_text(block), "Heading1"),
        "list" => block
            .get("items")
            .and_then(Value::as_array)
            .map(|items| {
                items
                    .iter()
                    .filter_map(Value::as_str)
                    .map(|item| docx_paragraph(&format!("• {item}"), "Normal"))
                    .collect::<String>()
            })
            .unwrap_or_default(),
        "table" => docx_table_xml(&block_rows(block)),
        "image" => docx_paragraph(&format!("[Image: {}]", alt_text(block)), "Normal"),
        "chart" => docx_paragraph(&format!("[Chart: {}]", block_text(block)), "Normal"),
        _ => docx_paragraph(&block_text(block), "Normal"),
    }
}

fn docx_paragraph(text: &str, style: &str) -> String {
    let style_xml = if style == "Normal" {
        String::new()
    } else {
        format!("<w:pPr><w:pStyle w:val=\"{}\"/></w:pPr>", xml_escape(style))
    };
    format!(
        "<w:p>{style_xml}<w:r><w:t>{}</w:t></w:r></w:p>",
        xml_escape(text)
    )
}

fn docx_table_xml(rows: &[Vec<String>]) -> String {
    let mut out = String::from("<w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/></w:tblPr>");
    for row in rows {
        out.push_str("<w:tr>");
        for cell in row {
            out.push_str(&format!(
                "<w:tc><w:tcPr><w:tcW w:w=\"2400\" w:type=\"dxa\"/></w:tcPr><w:p><w:r><w:t>{}</w:t></w:r></w:p></w:tc>",
                xml_escape(cell)
            ));
        }
        out.push_str("</w:tr>");
    }
    out.push_str("</w:tbl>");
    out
}

fn pptx_slide_xml(slide_number: usize, slide: &SlideSpec) -> String {
    let mut shapes = String::new();
    shapes.push_str(&pptx_text_shape(
        2,
        "Title",
        457200,
        274320,
        8229600,
        731520,
        &slide.title,
        3200,
    ));
    let mut next_id = 3usize;
    let mut y = 1219200i64;
    for block in &slide.blocks {
        let kind = block
            .get("type")
            .and_then(Value::as_str)
            .unwrap_or("paragraph");
        match kind {
            "table" => {
                shapes.push_str(&pptx_table_shape(next_id, y, &block_rows(block)));
                y += 1524000;
            }
            "list" => {
                let text = block
                    .get("items")
                    .and_then(Value::as_array)
                    .map(|items| {
                        items
                            .iter()
                            .filter_map(Value::as_str)
                            .map(|item| format!("• {item}"))
                            .collect::<Vec<_>>()
                            .join("\n")
                    })
                    .unwrap_or_default();
                shapes.push_str(&pptx_text_shape(
                    next_id, "List", 609600, y, 7924800, 1066800, &text, 2000,
                ));
                y += 1219200;
            }
            "image" => {
                shapes.push_str(&pptx_text_shape(
                    next_id,
                    "Image",
                    609600,
                    y,
                    3657600,
                    914400,
                    &format!("[Image: {}]", alt_text(block)),
                    1800,
                ));
                y += 1066800;
            }
            "chart" => {
                shapes.push_str(&pptx_text_shape(
                    next_id,
                    "Chart",
                    609600,
                    y,
                    3657600,
                    914400,
                    &format!("[Chart: {}]", block_text(block)),
                    1800,
                ));
                y += 1066800;
            }
            "heading" | "title" => {
                shapes.push_str(&pptx_text_shape(
                    next_id,
                    "Heading",
                    609600,
                    y,
                    7924800,
                    609600,
                    &block_text(block),
                    2400,
                ));
                y += 731520;
            }
            _ => {
                shapes.push_str(&pptx_text_shape(
                    next_id,
                    "Text",
                    609600,
                    y,
                    7924800,
                    914400,
                    &block_text(block),
                    1900,
                ));
                y += 1066800;
            }
        }
        next_id += 1;
    }
    format!("<p:sld xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><p:cSld name=\"Slide {slide_number}\"><p:spTree><p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/><a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"0\" cy=\"0\"/></a:xfrm></p:grpSpPr>{shapes}</p:spTree></p:cSld></p:sld>")
}

fn pptx_text_shape(
    id: usize,
    name: &str,
    x: i64,
    y: i64,
    w: i64,
    h: i64,
    text: &str,
    size: usize,
) -> String {
    let paragraphs = text
        .lines()
        .map(|line| {
            format!(
                "<a:p><a:r><a:rPr sz=\"{size}\"/><a:t>{}</a:t></a:r></a:p>",
                xml_escape(line)
            )
        })
        .collect::<String>();
    format!("<p:sp><p:nvSpPr><p:cNvPr id=\"{id}\" name=\"{}\"/><p:cNvSpPr txBox=\"1\"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x=\"{x}\" y=\"{y}\"/><a:ext cx=\"{w}\" cy=\"{h}\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr wrap=\"square\"/><a:lstStyle/>{paragraphs}</p:txBody></p:sp>", xml_escape(name))
}

fn pptx_table_shape(id: usize, y: i64, rows: &[Vec<String>]) -> String {
    let cols = rows.iter().map(Vec::len).max().unwrap_or(1).max(1);
    let mut grid = String::new();
    for _ in 0..cols {
        grid.push_str("<a:gridCol w=\"2200000\"/>");
    }
    let mut row_xml = String::new();
    for row in rows {
        row_xml.push_str("<a:tr h=\"370000\">");
        for col in 0..cols {
            let cell = row.get(col).cloned().unwrap_or_default();
            row_xml.push_str(&format!("<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>{}</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>", xml_escape(&cell)));
        }
        row_xml.push_str("</a:tr>");
    }
    format!("<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id=\"{id}\" name=\"Table\"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x=\"609600\" y=\"{y}\"/><a:ext cx=\"7924800\" cy=\"1600200\"/></p:xfrm><a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/table\"><a:tbl><a:tblPr/><a:tblGrid>{grid}</a:tblGrid>{row_xml}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>")
}

fn summarize_ooxml(path: &Path, format: &str) -> Result<OoxmlSummary> {
    let mut zip = ZipArchive::new(File::open(path)?)?;
    let mut parts = Vec::new();
    let mut texts = Vec::new();
    let mut object_map = Vec::new();
    let mut tables = Vec::new();
    let mut images = 0usize;
    let mut charts = 0usize;
    for i in 0..zip.len() {
        let mut file = zip.by_index(i)?;
        let name = file.name().to_string();
        parts.push(name.clone());
        if name.contains("/media/") {
            images += 1;
        }
        if name.contains("/charts/") {
            charts += 1;
        }
        if !name.ends_with(".xml") {
            continue;
        }
        let mut xml = String::new();
        if file.read_to_string(&mut xml).is_err() {
            continue;
        }
        let semantic_part = format != "pptx" || name.starts_with("ppt/slides/slide");
        if semantic_part {
            images += xml.matches("<a:blip").count();
            charts += xml.matches("<c:chart").count();
            images += xml.matches("[Image:").count();
            charts += xml.matches("[Chart:").count();
        }

        if !semantic_part {
            continue;
        }

        for (idx, text) in xml_text_nodes(&xml).iter().enumerate() {
            if text.trim().is_empty() {
                continue;
            }
            let id = stable_id(format, &name, idx, "");
            object_map.push(json!({
                "stableObjectId": id,
                "type": "text",
                "sourcePath": name,
                "textPreview": text,
                "selectorHints": {"contains": text, "stableObjectId": id, "sourcePath": name}
            }));
            texts.push(text.clone());
        }
        for (idx, rows) in table_nodes(format, &xml).into_iter().enumerate() {
            let preview = rows
                .iter()
                .flat_map(|row| row.iter())
                .cloned()
                .collect::<Vec<_>>()
                .join(" ");
            object_map.push(json!({
                "stableObjectId": stable_id(format, &format!("{name}:table"), idx, ""),
                "type": "table",
                "sourcePath": name,
                "rowCount": rows.len(),
                "columnCount": rows.iter().map(Vec::len).max().unwrap_or(0),
                "rows": rows,
                "textPreview": preview,
                "selectorHints": {"tableIndex": idx, "sourcePath": name}
            }));
            tables.push(rows);
        }
    }
    Ok(OoxmlSummary {
        format: format.to_string(),
        parts,
        texts,
        tables,
        images,
        charts,
        object_map,
    })
}

fn table_nodes(format: &str, xml: &str) -> Vec<Vec<Vec<String>>> {
    let (table_tag, row_tag, cell_tag) = if format == "pptx" {
        ("a:tbl", "a:tr", "a:tc")
    } else {
        ("w:tbl", "w:tr", "w:tc")
    };
    let table_re = tag_regex(table_tag);
    let row_re = tag_regex(row_tag);
    let cell_re = tag_regex(cell_tag);
    let mut out = Vec::new();
    for table_cap in table_re.captures_iter(xml) {
        let Some(table_xml) = table_cap.get(0).map(|m| m.as_str()) else {
            continue;
        };
        let mut rows = Vec::new();
        for row_cap in row_re.captures_iter(table_xml) {
            let Some(row_xml) = row_cap.get(0).map(|m| m.as_str()) else {
                continue;
            };
            let row = cell_re
                .captures_iter(row_xml)
                .filter_map(|cell_cap| cell_cap.get(0))
                .map(|cell| xml_text_nodes(cell.as_str()).join(" "))
                .collect::<Vec<_>>();
            if !row.is_empty() {
                rows.push(row);
            }
        }
        if !rows.is_empty() {
            out.push(rows);
        }
    }
    out
}

fn replace_table_cell(
    xml: &str,
    table_index: usize,
    row_index: usize,
    col_index: usize,
    replacement: &str,
    table_tag: &str,
    row_tag: &str,
    cell_tag: &str,
    text_tag: &str,
) -> Result<Option<String>> {
    let table_re = tag_regex(table_tag);
    let row_re = tag_regex(row_tag);
    let cell_re = tag_regex(cell_tag);
    let text_re = tag_regex(text_tag);
    let Some(table_match) = table_re.find_iter(xml).nth(table_index) else {
        return Ok(None);
    };
    let table_xml = table_match.as_str();
    let Some(row_match) = row_re.find_iter(table_xml).nth(row_index) else {
        bail!("SELECTOR_NOT_FOUND: table row was not found");
    };
    let row_xml = row_match.as_str();
    let Some(cell_match) = cell_re.find_iter(row_xml).nth(col_index) else {
        bail!("SELECTOR_NOT_FOUND: table cell was not found");
    };
    let cell_xml = cell_match.as_str();
    let Some(text_match) = text_re.find_iter(cell_xml).next() else {
        bail!("SELECTOR_NOT_FOUND: table cell has no editable text node");
    };
    let global_start =
        table_match.start() + row_match.start() + cell_match.start() + text_match.start();
    let global_end =
        table_match.start() + row_match.start() + cell_match.start() + text_match.end();
    let original_text_node = &xml[global_start..global_end];
    let open_end = original_text_node
        .find('>')
        .ok_or_else(|| anyhow!("OOXML_PARSE_FAILED: text node is malformed"))?
        + 1;
    let close_start = original_text_node
        .rfind("</")
        .ok_or_else(|| anyhow!("OOXML_PARSE_FAILED: text node is malformed"))?;
    let mut replacement_node = String::new();
    replacement_node.push_str(&original_text_node[..open_end]);
    replacement_node.push_str(&xml_escape(replacement));
    replacement_node.push_str(&original_text_node[close_start..]);
    let mut next = String::with_capacity(xml.len() + replacement.len());
    next.push_str(&xml[..global_start]);
    next.push_str(&replacement_node);
    next.push_str(&xml[global_end..]);
    Ok(Some(next))
}

fn tag_regex(tag: &str) -> Regex {
    let close = regex::escape(tag);
    let local = tag.split(':').last().unwrap_or(tag);
    let open = if let Some(prefix) = tag.split(':').next().filter(|prefix| *prefix != local) {
        format!(
            r"(?:{}|{})",
            regex::escape(tag),
            regex::escape(local).replace(local, &format!("{prefix}:{local}"))
        )
    } else {
        regex::escape(tag)
    };
    Regex::new(&format!(r"(?s)<{open}(?:\s[^>]*)?>.*?</{close}>")).unwrap()
}

fn index_arg(op: &Value, zero_based: &str, one_based: &str) -> Result<usize> {
    if let Some(value) = op
        .pointer(&format!("/selector/{zero_based}"))
        .and_then(Value::as_u64)
        .or_else(|| op.get(zero_based).and_then(Value::as_u64))
    {
        return Ok(value as usize);
    }
    if let Some(value) = op
        .pointer(&format!("/selector/{one_based}"))
        .and_then(Value::as_u64)
        .or_else(|| op.get(one_based).and_then(Value::as_u64))
    {
        return Ok(value.saturating_sub(1) as usize);
    }
    bail!("SCHEMA_INVALID: table edit requires selector.{zero_based} or selector.{one_based}");
}

fn block_rows(block: &Value) -> Vec<Vec<String>> {
    block
        .get("rows")
        .or_else(|| block.get("data"))
        .and_then(Value::as_array)
        .map(|rows| {
            rows.iter()
                .map(|row| {
                    row.as_array()
                        .map(|cells| cells.iter().map(cell_to_string).collect::<Vec<_>>())
                        .unwrap_or_else(|| vec![cell_to_string(row)])
                })
                .collect::<Vec<_>>()
        })
        .filter(|rows| !rows.is_empty())
        .unwrap_or_else(|| vec![vec!["Metric".into(), "Value".into()]])
}

fn block_text(block: &Value) -> String {
    for key in ["text", "title", "content", "caption", "label"] {
        if let Some(text) = block.get(key).and_then(Value::as_str) {
            return text.to_string();
        }
    }
    collect_ir_text(block)
}

fn alt_text(block: &Value) -> String {
    block
        .get("alt")
        .or_else(|| block.get("description"))
        .or_else(|| block.get("title"))
        .and_then(Value::as_str)
        .unwrap_or("image")
        .to_string()
}

fn cell_to_string(value: &Value) -> String {
    match value {
        Value::String(s) => s.clone(),
        Value::Number(n) => n.to_string(),
        Value::Bool(b) => b.to_string(),
        _ => value.to_string(),
    }
}

fn collect_ir_text(value: &Value) -> String {
    let mut out = Vec::new();
    collect_strings(value, &mut out);
    out.join(" ")
}

fn collect_strings(value: &Value, out: &mut Vec<String>) {
    match value {
        Value::String(s) => out.push(s.clone()),
        Value::Array(items) => items.iter().for_each(|v| collect_strings(v, out)),
        Value::Object(map) => map.values().for_each(|v| collect_strings(v, out)),
        _ => {}
    }
}

fn xml_text_nodes(xml: &str) -> Vec<String> {
    let re = Regex::new(r"(?s)<(?:a:t|w:t|t|v)(?:\s[^>]*)?>(.*?)</(?:a:t|w:t|t|v)>").unwrap();
    re.captures_iter(xml)
        .filter_map(|cap| cap.get(1))
        .map(|m| xml_unescape(m.as_str()))
        .collect()
}

fn xml_escape(text: &str) -> String {
    text.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
}

fn xml_unescape(text: &str) -> String {
    text.replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&quot;", "\"")
        .replace("&amp;", "&")
}

fn html_escape(text: &str) -> String {
    xml_escape(text)
}

fn stable_id(format: &str, part: &str, idx: usize, _text: &str) -> String {
    let mut hash = Sha256::new();
    hash.update(format.as_bytes());
    hash.update(part.as_bytes());
    hash.update(idx.to_string().as_bytes());
    format!("{format}:{}", &hex::encode(hash.finalize())[..16])
}

fn sha256_file(path: &Path) -> Result<String> {
    let mut hash = Sha256::new();
    hash.update(fs::read(path)?);
    Ok(format!("sha256:{}", hex::encode(hash.finalize())))
}

fn table_fingerprint(tables: &[Vec<Vec<String>>]) -> String {
    serde_json::to_string(tables).unwrap_or_default()
}

fn presence_json(summary: &OoxmlSummary) -> Value {
    json!({
        "format": summary.format,
        "text": summary.texts.iter().any(|text| !text.trim().is_empty()),
        "tables": summary.tables.len(),
        "images": summary.images,
        "charts": summary.charts
    })
}

fn summary_chips(summary: &Value) -> String {
    let mut chips = Vec::new();
    for key in [
        "format",
        "textObjects",
        "tables",
        "images",
        "charts",
        "parts",
    ] {
        if let Some(value) = summary.get(key) {
            chips.push(format!(
                "<span class=\"chip\">{}: {}</span>",
                html_escape(key),
                html_escape(&value.to_string().trim_matches('"').to_string())
            ));
        }
    }
    chips.join("")
}

fn table_html(object: &Value) -> String {
    let Some(rows) = object.get("rows").and_then(Value::as_array) else {
        return String::new();
    };
    let mut out = String::from("<table>");
    for row in rows {
        out.push_str("<tr>");
        for cell in row.as_array().into_iter().flatten() {
            out.push_str(&format!(
                "<td>{}</td>",
                html_escape(cell.as_str().unwrap_or(""))
            ));
        }
        out.push_str("</tr>");
    }
    out.push_str("</table>");
    out
}

fn zip_file<W: Write + std::io::Seek>(
    writer: &mut ZipWriter<W>,
    name: &str,
    content: impl AsRef<[u8]>,
    options: SimpleFileOptions,
) -> Result<()> {
    writer.start_file(name, options)?;
    writer.write_all(content.as_ref())?;
    Ok(())
}

fn rels(rel_type: &str, target: &str) -> String {
    format!(
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/{rel_type}" Target="{target}"/></Relationships>"#
    )
}

fn pptx_content_types(slides: usize) -> String {
    let mut overrides = String::from(
        r#"<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>"#,
    );
    for index in 1..=slides {
        overrides.push_str(&format!(r#"<Override PartName="/ppt/slides/slide{index}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>"#));
    }
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>{overrides}</Types>"#
    )
}

fn docx_content_types() -> String {
    r#"<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_string()
}

fn pptx_presentation_xml(slides: usize) -> String {
    let mut slide_ids = String::new();
    for index in 1..=slides {
        slide_ids.push_str(&format!(
            "<p:sldId id=\"{}\" r:id=\"rId{}\"/>",
            255 + index,
            index
        ));
    }
    format!("<p:presentation xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><p:sldIdLst>{slide_ids}</p:sldIdLst></p:presentation>")
}

fn pptx_presentation_rels(slides: usize) -> String {
    let mut rels_xml = String::from(
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
    );
    for index in 1..=slides {
        rels_xml.push_str(&format!(r#"<Relationship Id="rId{index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{index}.xml"/>"#));
    }
    rels_xml.push_str("</Relationships>");
    rels_xml
}
