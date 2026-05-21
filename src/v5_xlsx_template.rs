use crate::safety;
use anyhow::{anyhow, bail, Result};
use regex::Regex;
use serde_json::{json, Value};
use sha2::{Digest, Sha256};
use std::collections::BTreeMap;
use std::io::{Cursor, Write};
use std::path::Path;
use zip::write::SimpleFileOptions;
use zip::ZipWriter;

#[derive(Clone, Debug)]
struct XlsxSheet {
    name: String,
    cells: BTreeMap<String, XlsxCell>,
    tables: Vec<Value>,
    charts: Vec<Value>,
    validations: Vec<Value>,
}

#[derive(Clone, Debug)]
enum XlsxCell {
    Value(Value),
    Formula(String),
}

pub fn write_xlsx_from_ir(path: &Path, ir: &Value, title: &str, text: &str) -> Result<()> {
    let mut sheets = xlsx_sheets_from_ir(ir)?;
    if sheets.is_empty() {
        let mut cells = BTreeMap::new();
        cells.insert("A1".into(), XlsxCell::Value(json!(title)));
        cells.insert("A2".into(), XlsxCell::Value(json!(text)));
        sheets.push(XlsxSheet {
            name: "Sheet1".into(),
            cells,
            tables: Vec::new(),
            charts: Vec::new(),
            validations: Vec::new(),
        });
    }
    write_xlsx_package(path, &sheets)
}

pub fn apply_xlsx_package_op(entries: &mut Vec<(String, Vec<u8>)>, op: &Value) -> Result<bool> {
    let op_name = op
        .get("op")
        .or_else(|| op.get("type"))
        .and_then(Value::as_str)
        .unwrap_or("");
    match op_name {
        "xlsx.setCell" => {
            let value = op
                .get("value")
                .or_else(|| op.get("text"))
                .ok_or_else(|| anyhow!("SCHEMA_INVALID: xlsx.setCell requires value"))?;
            let (part, cell) = worksheet_cell_for_op(entries, op)?;
            let idx = entry_index(entries, &part)
                .ok_or_else(|| anyhow!("SELECTOR_NOT_FOUND: worksheet {part} was not found"))?;
            let xml = entry_text(entries, idx)?;
            let mut cells = BTreeMap::new();
            cells.insert(cell.to_ascii_uppercase(), XlsxCell::Value(value.clone()));
            let next = set_cells_in_sheet_xml(&xml, &cells);
            if next == xml {
                return Ok(false);
            }
            entries[idx].1 = next.into_bytes();
            Ok(true)
        }
        "xlsx.setFormula" => {
            let formula = op
                .get("formula")
                .and_then(Value::as_str)
                .ok_or_else(|| anyhow!("SCHEMA_INVALID: xlsx.setFormula requires formula"))?;
            validate_formula_safety(formula)?;
            let (part, cell) = worksheet_cell_for_op(entries, op)?;
            let idx = entry_index(entries, &part)
                .ok_or_else(|| anyhow!("SELECTOR_NOT_FOUND: worksheet {part} was not found"))?;
            let xml = entry_text(entries, idx)?;
            let next = set_xlsx_cell_xml(&xml, &cell, formula, true);
            if next == xml {
                return Ok(false);
            }
            entries[idx].1 = next.into_bytes();
            Ok(true)
        }
        "xlsx.setRange" => {
            let values = op
                .get("values")
                .or_else(|| op.get("rows"))
                .and_then(Value::as_array)
                .ok_or_else(|| anyhow!("SCHEMA_INVALID: xlsx.setRange requires values array"))?;
            let start = op
                .get("cell")
                .or_else(|| op.get("startCell"))
                .or_else(|| op.pointer("/selector/cell"))
                .and_then(Value::as_str)
                .unwrap_or("A1");
            let part = worksheet_part_for_op(entries, op)?;
            let idx = entry_index(entries, &part)
                .ok_or_else(|| anyhow!("SELECTOR_NOT_FOUND: worksheet {part} was not found"))?;
            let xml = entry_text(entries, idx)?;
            let cells = range_cells(start, values)?;
            let next = set_cells_in_sheet_xml(&xml, &cells);
            if next == xml {
                return Ok(false);
            }
            entries[idx].1 = next.into_bytes();
            Ok(true)
        }
        "xlsx.addSheet" => add_sheet(entries, op),
        "xlsx.renameSheet" => rename_sheet(entries, op),
        "xlsx.addTable" => add_table(entries, op),
        "xlsx.setNamedRange" => set_named_range(entries, op),
        "xlsx.setDataValidation" => set_data_validation(entries, op),
        "xlsx.addChart" => add_chart(entries, op),
        _ => Ok(false),
    }
}

pub fn set_xlsx_cell_xml(xml: &str, cell: &str, value: &str, formula: bool) -> String {
    let mut cells = BTreeMap::new();
    let xlsx_cell = if formula {
        XlsxCell::Formula(value.to_string())
    } else {
        XlsxCell::Value(json!(value))
    };
    cells.insert(cell.to_ascii_uppercase(), xlsx_cell);
    set_cells_in_sheet_xml(xml, &cells)
}

pub fn validate_formula_safety(formula: &str) -> Result<()> {
    let normalized = formula.trim().trim_start_matches('=').to_ascii_lowercase();
    let risky_tokens = [
        "webservice(",
        "hyperlink(",
        "cmd|",
        "powershell|",
        "dde(",
        "shell(",
        "http://",
        "https://",
        "ftp://",
        "file://",
        "]!",
    ];
    if risky_tokens.iter().any(|token| normalized.contains(token)) {
        bail!("SCHEMA_INVALID: xlsx formulas cannot contain external links, network calls, or command-style references");
    }
    let external_ref = Regex::new(r"(?i)\[[^\]]+\][^!]*!").unwrap();
    if external_ref.is_match(formula) {
        bail!("SCHEMA_INVALID: xlsx formulas cannot contain external workbook references");
    }
    Ok(())
}

fn xlsx_sheets_from_ir(ir: &Value) -> Result<Vec<XlsxSheet>> {
    let mut out = Vec::new();
    if let Some(sheets) = ir.get("sheets").and_then(Value::as_array) {
        for (idx, sheet) in sheets.iter().enumerate() {
            out.push(xlsx_sheet_from_value(sheet, idx)?);
        }
    }
    if out.is_empty() {
        let blocks = ir
            .get("sections")
            .and_then(Value::as_array)
            .into_iter()
            .flatten()
            .flat_map(|section| {
                section
                    .get("blocks")
                    .and_then(Value::as_array)
                    .into_iter()
                    .flatten()
            });
        for block in blocks {
            if block.get("type").and_then(Value::as_str) == Some("table") {
                let mut sheet = xlsx_sheet_from_value(block, out.len())?;
                if sheet.name == format!("Sheet{}", out.len() + 1) {
                    sheet.name = block
                        .get("title")
                        .and_then(Value::as_str)
                        .unwrap_or("Sheet1")
                        .to_string();
                }
                out.push(sheet);
            }
        }
    }
    Ok(out)
}

fn xlsx_sheet_from_value(value: &Value, idx: usize) -> Result<XlsxSheet> {
    let name = value
        .get("name")
        .or_else(|| value.get("title"))
        .and_then(Value::as_str)
        .map(valid_sheet_name)
        .unwrap_or_else(|| format!("Sheet{}", idx + 1));
    let mut cells = BTreeMap::new();
    if let Some(rows) = value.get("rows").and_then(Value::as_array) {
        add_rows_to_cells(&mut cells, "A1", rows)?;
    }
    if let Some(cells_obj) = value.get("cells").and_then(Value::as_object) {
        for (cell, v) in cells_obj {
            cells.insert(cell.to_ascii_uppercase(), cell_from_value(v)?);
        }
    }
    if let Some(formulas) = value.get("formulas").and_then(Value::as_array) {
        for formula in formulas {
            if let (Some(cell), Some(expr)) = (
                formula.get("cell").and_then(Value::as_str),
                formula.get("formula").and_then(Value::as_str),
            ) {
                validate_formula_safety(expr)?;
                cells.insert(
                    cell.to_ascii_uppercase(),
                    XlsxCell::Formula(expr.to_string()),
                );
            }
        }
    }
    if value
        .get("charts")
        .and_then(Value::as_array)
        .map(|charts| charts.len() > 1)
        .unwrap_or(false)
    {
        bail!("FEATURE_NOT_IMPLEMENTED: XLSX render currently supports one chart per sheet");
    }
    Ok(XlsxSheet {
        name,
        cells,
        tables: value
            .get("tables")
            .and_then(Value::as_array)
            .cloned()
            .unwrap_or_default(),
        charts: value
            .get("charts")
            .and_then(Value::as_array)
            .cloned()
            .unwrap_or_default(),
        validations: value
            .get("validations")
            .or_else(|| value.get("dataValidations"))
            .and_then(Value::as_array)
            .cloned()
            .unwrap_or_default(),
    })
}

fn write_xlsx_package(path: &Path, sheets: &[XlsxSheet]) -> Result<()> {
    let mut entries = Vec::new();
    entries.push((
        "[Content_Types].xml".into(),
        xlsx_content_types(sheets).into_bytes(),
    ));
    entries.push(("_rels/.rels".into(), xlsx_root_rels().into_bytes()));
    entries.push(("xl/workbook.xml".into(), workbook_xml(sheets).into_bytes()));
    entries.push((
        "xl/_rels/workbook.xml.rels".into(),
        workbook_rels_xml(sheets.len()).into_bytes(),
    ));
    for (idx, sheet) in sheets.iter().enumerate() {
        entries.push((
            format!("xl/worksheets/sheet{}.xml", idx + 1),
            worksheet_xml(sheets, idx).into_bytes(),
        ));
        let rels = worksheet_rels_xml(sheets, idx);
        if !rels.is_empty() {
            entries.push((
                format!("xl/worksheets/_rels/sheet{}.xml.rels", idx + 1),
                rels.into_bytes(),
            ));
        }
        for (table_idx, table) in sheet.tables.iter().enumerate() {
            let id = table_part_id(sheets, idx, table_idx);
            entries.push((
                format!("xl/tables/table{id}.xml"),
                table_xml(table, id).into_bytes(),
            ));
        }
        for (chart_idx, chart) in sheet.charts.iter().enumerate() {
            let chart_id = chart_part_id(sheets, idx, chart_idx);
            entries.push((
                format!("xl/drawings/drawing{chart_id}.xml"),
                drawing_xml(chart_id).into_bytes(),
            ));
            entries.push((
                format!("xl/drawings/_rels/drawing{chart_id}.xml.rels"),
                drawing_rels_xml(chart_id).into_bytes(),
            ));
            entries.push((
                format!("xl/charts/chart{chart_id}.xml"),
                chart_xml(chart, &format!("Chart {chart_id}")).into_bytes(),
            ));
        }
    }
    let bytes = zip_entries_to_bytes(entries)?;
    safety::atomic_write(path, &bytes)?;
    Ok(())
}

fn xlsx_content_types(sheets: &[XlsxSheet]) -> String {
    let mut overrides = String::from(
        r#"<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>"#,
    );
    for (idx, sheet) in sheets.iter().enumerate() {
        overrides.push_str(&format!(r#"<Override PartName="/xl/worksheets/sheet{}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>"#, idx + 1));
        for table_idx in 0..sheet.tables.len() {
            let id = table_part_id(sheets, idx, table_idx);
            overrides.push_str(&format!(r#"<Override PartName="/xl/tables/table{id}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>"#));
        }
        for chart_idx in 0..sheet.charts.len() {
            let id = chart_part_id(sheets, idx, chart_idx);
            overrides.push_str(&format!(r#"<Override PartName="/xl/drawings/drawing{id}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>"#));
            overrides.push_str(&format!(r#"<Override PartName="/xl/charts/chart{id}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>"#));
        }
    }
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>{overrides}</Types>"#
    )
}

fn xlsx_root_rels() -> String {
    r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#.into()
}

fn workbook_xml(sheets: &[XlsxSheet]) -> String {
    let sheet_xml = sheets
        .iter()
        .enumerate()
        .map(|(idx, sheet)| {
            format!(
                r#"<sheet name="{}" sheetId="{}" r:id="rId{}"/>"#,
                xml_escape(&sheet.name),
                idx + 1,
                idx + 1
            )
        })
        .collect::<String>();
    format!(
        r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>{sheet_xml}</sheets></workbook>"#
    )
}

fn workbook_rels_xml(sheet_count: usize) -> String {
    let rels = (1..=sheet_count)
        .map(|idx| format!(r#"<Relationship Id="rId{idx}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{idx}.xml"/>"#))
        .collect::<String>();
    format!(
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{rels}</Relationships>"#
    )
}

fn worksheet_xml(sheets: &[XlsxSheet], sheet_idx: usize) -> String {
    let sheet = &sheets[sheet_idx];
    let mut rows: BTreeMap<usize, Vec<(String, XlsxCell)>> = BTreeMap::new();
    for (cell, value) in &sheet.cells {
        let (_, row) = split_cell_ref(cell).unwrap_or((1, 1));
        rows.entry(row)
            .or_default()
            .push((cell.clone(), value.clone()));
    }
    let row_xml = rows
        .into_iter()
        .map(|(row, mut cells)| {
            cells.sort_by_key(|(cell, _)| split_cell_ref(cell).map(|(col, _)| col).unwrap_or(1));
            let cells_xml = cells
                .into_iter()
                .map(|(cell, value)| cell_xml(&cell, &value))
                .collect::<String>();
            format!(r#"<row r="{row}">{cells_xml}</row>"#)
        })
        .collect::<String>();
    let validations = validations_xml(&sheet.validations);
    let drawing = if sheet.charts.is_empty() {
        String::new()
    } else {
        format!(r#"<drawing r:id="rId{}"/>"#, drawing_rel_id(sheet))
    };
    let table_parts = if sheet.tables.is_empty() {
        String::new()
    } else {
        let parts = sheet
            .tables
            .iter()
            .enumerate()
            .map(|(idx, _)| format!(r#"<tablePart r:id="rId{}"/>"#, table_rel_id(idx)))
            .collect::<String>();
        format!(
            r#"<tableParts count="{}">{parts}</tableParts>"#,
            sheet.tables.len()
        )
    };
    format!(
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData>{row_xml}</sheetData>{validations}{drawing}{table_parts}</worksheet>"#
    )
}

fn worksheet_rels_xml(sheets: &[XlsxSheet], sheet_idx: usize) -> String {
    let sheet = &sheets[sheet_idx];
    if sheet.tables.is_empty() && sheet.charts.is_empty() {
        return String::new();
    }
    let mut rels = String::new();
    for (idx, _) in sheet.tables.iter().enumerate() {
        let rel_id = table_rel_id(idx);
        let part_id = table_part_id(sheets, sheet_idx, idx);
        rels.push_str(&format!(r#"<Relationship Id="rId{rel_id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table{part_id}.xml"/>"#));
    }
    if !sheet.charts.is_empty() {
        let rel_id = drawing_rel_id(sheet);
        let chart_id = chart_part_id(sheets, sheet_idx, 0);
        rels.push_str(&format!(r#"<Relationship Id="rId{rel_id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing{chart_id}.xml"/>"#));
    }
    format!(
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{rels}</Relationships>"#
    )
}

fn table_xml(table: &Value, id: usize) -> String {
    let name = table
        .get("name")
        .and_then(Value::as_str)
        .map(valid_table_name)
        .unwrap_or_else(|| format!("Table{id}"));
    let range = table
        .get("range")
        .or_else(|| table.pointer("/selector/range"))
        .and_then(Value::as_str)
        .unwrap_or("A1:B2");
    let column_count = table_column_count(range).unwrap_or(1);
    let columns = (1..=column_count)
        .map(|idx| format!(r#"<tableColumn id="{idx}" name="Column{idx}"/>"#))
        .collect::<String>();
    format!(
        r#"<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="{id}" name="{name}" displayName="{name}" ref="{range}" totalsRowShown="0"><autoFilter ref="{range}"/><tableColumns count="{column_count}">{columns}</tableColumns><tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>"#
    )
}

fn drawing_xml(chart_id: usize) -> String {
    format!(
        r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><xdr:oneCellAnchor><xdr:from><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="5486400" cy="3200400"/><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart {chart_id}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:oneCellAnchor></xdr:wsDr>"#
    )
}

fn drawing_rels_xml(chart_id: usize) -> String {
    format!(
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart{chart_id}.xml"/></Relationships>"#
    )
}

fn chart_xml(chart: &Value, fallback_title: &str) -> String {
    let (labels, values) = chart_data(chart).unwrap_or_else(|_| (vec!["Value".into()], vec![0.0]));
    let title = chart
        .get("title")
        .and_then(Value::as_str)
        .unwrap_or(fallback_title);
    let label_points = labels
        .iter()
        .enumerate()
        .map(|(idx, label)| {
            format!(
                r#"<c:pt idx="{idx}"><c:v>{}</c:v></c:pt>"#,
                xml_escape(label)
            )
        })
        .collect::<String>();
    let value_points = values
        .iter()
        .enumerate()
        .map(|(idx, value)| format!(r#"<c:pt idx="{idx}"><c:v>{value}</c:v></c:pt>"#))
        .collect::<String>();
    let count = labels.len().min(values.len());
    format!(
        r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>{}</a:t></a:r></a:p></c:rich></c:tx></c:title><c:plotArea><c:barChart><c:barDir val="col"/><c:ser><c:idx val="0"/><c:order val="0"/><c:cat><c:strLit><c:ptCount val="{count}"/>{label_points}</c:strLit></c:cat><c:val><c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="{count}"/>{value_points}</c:numLit></c:val></c:ser></c:barChart></c:plotArea><c:legend><c:legendPos val="r"/></c:legend><c:plotVisOnly val="1"/></c:chart></c:chartSpace>"#,
        xml_escape(title)
    )
}

fn add_sheet(entries: &mut Vec<(String, Vec<u8>)>, op: &Value) -> Result<bool> {
    let name = op
        .get("name")
        .or_else(|| op.get("sheet"))
        .and_then(Value::as_str)
        .map(valid_sheet_name)
        .ok_or_else(|| anyhow!("SCHEMA_INVALID: xlsx.addSheet requires name"))?;
    let workbook_idx = entry_index(entries, "xl/workbook.xml")
        .ok_or_else(|| anyhow!("SELECTOR_NOT_FOUND: xl/workbook.xml"))?;
    let rels_idx = entry_index(entries, "xl/_rels/workbook.xml.rels")
        .ok_or_else(|| anyhow!("SELECTOR_NOT_FOUND: xl/_rels/workbook.xml.rels"))?;
    let workbook = entry_text(entries, workbook_idx)?;
    if workbook.contains(&format!(r#"name="{}""#, xml_escape(&name))) {
        bail!("SCHEMA_INVALID: xlsx.addSheet sheet already exists");
    }
    let next_sheet = worksheet_count(entries) + 1;
    let next_rel = next_rid(&entry_text(entries, rels_idx)?);
    let sheet = format!(
        r#"<sheet name="{}" sheetId="{next_sheet}" r:id="rId{next_rel}"/>"#,
        xml_escape(&name)
    );
    let next_workbook = insert_before(&workbook, "</sheets>", &sheet)?;
    let rel = format!(
        r#"<Relationship Id="rId{next_rel}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{next_sheet}.xml"/>"#
    );
    let rels = entry_text(entries, rels_idx)?;
    let next_rels = insert_before(&rels, "</Relationships>", &rel)?;
    entries[workbook_idx].1 = next_workbook.into_bytes();
    entries[rels_idx].1 = next_rels.into_bytes();
    entries.push((
        format!("xl/worksheets/sheet{next_sheet}.xml"),
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>"#
            .as_bytes()
            .to_vec(),
    ));
    upsert_content_type(
        entries,
        &format!("/xl/worksheets/sheet{next_sheet}.xml"),
        "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
    )?;
    Ok(true)
}

fn rename_sheet(entries: &mut Vec<(String, Vec<u8>)>, op: &Value) -> Result<bool> {
    let from = op
        .get("sheet")
        .or_else(|| op.get("from"))
        .or_else(|| op.pointer("/selector/sheet"))
        .and_then(Value::as_str)
        .ok_or_else(|| anyhow!("SCHEMA_INVALID: xlsx.renameSheet requires sheet/from"))?;
    let to = op
        .get("name")
        .or_else(|| op.get("to"))
        .and_then(Value::as_str)
        .map(valid_sheet_name)
        .ok_or_else(|| anyhow!("SCHEMA_INVALID: xlsx.renameSheet requires name/to"))?;
    let workbook_idx = entry_index(entries, "xl/workbook.xml")
        .ok_or_else(|| anyhow!("SELECTOR_NOT_FOUND: xl/workbook.xml"))?;
    let workbook = entry_text(entries, workbook_idx)?;
    let re = Regex::new(&format!(
        r#"(<sheet\b[^>]*\bname="){}("[^>]*/>)"#,
        regex::escape(&xml_escape(from))
    ))?;
    if !re.is_match(&workbook) {
        return Ok(false);
    }
    entries[workbook_idx].1 = re
        .replace(&workbook, format!("${{1}}{}${{2}}", xml_escape(&to)))
        .to_string()
        .into_bytes();
    Ok(true)
}

fn add_table(entries: &mut Vec<(String, Vec<u8>)>, op: &Value) -> Result<bool> {
    let range = op
        .get("range")
        .or_else(|| op.pointer("/selector/range"))
        .and_then(Value::as_str)
        .ok_or_else(|| anyhow!("SCHEMA_INVALID: xlsx.addTable requires range"))?;
    let part = worksheet_part_for_op(entries, op)?;
    let sheet_idx = entry_index(entries, &part)
        .ok_or_else(|| anyhow!("SELECTOR_NOT_FOUND: worksheet {part} was not found"))?;
    let table_id = next_table_id(entries);
    let table =
        json!({"name": op.get("name").and_then(Value::as_str).unwrap_or("Table"), "range": range});
    let table_part = format!("xl/tables/table{table_id}.xml");
    entries.push((table_part, table_xml(&table, table_id).into_bytes()));
    upsert_content_type(
        entries,
        &format!("/xl/tables/table{table_id}.xml"),
        "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml",
    )?;
    let rel_id = add_sheet_relationship(
        entries,
        &part,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table",
        &format!("../tables/table{table_id}.xml"),
    )?;
    let sheet = entry_text(entries, sheet_idx)?;
    let table_part_xml = format!(r#"<tablePart r:id="rId{rel_id}"/>"#);
    let next = if sheet.contains("<tableParts") {
        insert_before(&sheet, "</tableParts>", &table_part_xml)?
    } else {
        insert_before(
            &ensure_worksheet_r_namespace(&sheet),
            "</worksheet>",
            &format!(r#"<tableParts count="1">{table_part_xml}</tableParts>"#),
        )?
    };
    entries[sheet_idx].1 = update_table_parts_count(&next).into_bytes();
    Ok(true)
}

fn set_named_range(entries: &mut Vec<(String, Vec<u8>)>, op: &Value) -> Result<bool> {
    let name = op
        .get("name")
        .and_then(Value::as_str)
        .ok_or_else(|| anyhow!("SCHEMA_INVALID: xlsx.setNamedRange requires name"))?;
    let refers_to = op
        .get("refersTo")
        .or_else(|| op.get("range"))
        .and_then(Value::as_str)
        .ok_or_else(|| anyhow!("SCHEMA_INVALID: xlsx.setNamedRange requires refersTo/range"))?;
    let workbook_idx = entry_index(entries, "xl/workbook.xml")
        .ok_or_else(|| anyhow!("SELECTOR_NOT_FOUND: xl/workbook.xml"))?;
    let workbook = entry_text(entries, workbook_idx)?;
    let defined_name = format!(
        r#"<definedName name="{}">{}</definedName>"#,
        xml_escape(name),
        xml_escape(refers_to)
    );
    let next = if workbook.contains("<definedNames>") {
        let re = Regex::new(&format!(
            r#"(?s)<definedName\s+[^>]*name="{}"[^>]*>.*?</definedName>"#,
            regex::escape(&xml_escape(name))
        ))?;
        if re.is_match(&workbook) {
            re.replace(&workbook, defined_name).to_string()
        } else {
            insert_before(&workbook, "</definedNames>", &defined_name)?
        }
    } else {
        insert_before(
            &workbook,
            "</workbook>",
            &format!("<definedNames>{defined_name}</definedNames>"),
        )?
    };
    entries[workbook_idx].1 = next.into_bytes();
    Ok(true)
}

fn set_data_validation(entries: &mut Vec<(String, Vec<u8>)>, op: &Value) -> Result<bool> {
    let range = op
        .get("range")
        .or_else(|| op.pointer("/selector/range"))
        .and_then(Value::as_str)
        .ok_or_else(|| anyhow!("SCHEMA_INVALID: xlsx.setDataValidation requires range"))?;
    let validation = json!({
        "range": range,
        "type": op.get("validationType").or_else(|| op.get("type")).and_then(Value::as_str).unwrap_or("list"),
        "formula1": op.get("formula1").or_else(|| op.get("source")).and_then(Value::as_str).unwrap_or("")
    });
    let part = worksheet_part_for_op(entries, op)?;
    let idx = entry_index(entries, &part)
        .ok_or_else(|| anyhow!("SELECTOR_NOT_FOUND: worksheet {part} was not found"))?;
    let sheet = entry_text(entries, idx)?;
    let validation_xml = data_validation_xml(&validation)
        .ok_or_else(|| anyhow!("SCHEMA_INVALID: xlsx.setDataValidation requires valid range"))?;
    let next = if sheet.contains("<dataValidations") {
        insert_before(&sheet, "</dataValidations>", &validation_xml)?
    } else {
        insert_before(
            &sheet,
            "</worksheet>",
            &format!(r#"<dataValidations count="1">{validation_xml}</dataValidations>"#),
        )?
    };
    entries[idx].1 = update_data_validations_count(&next).into_bytes();
    Ok(true)
}

fn add_chart(entries: &mut Vec<(String, Vec<u8>)>, op: &Value) -> Result<bool> {
    chart_data(op)?;
    let part = worksheet_part_for_op(entries, op)?;
    let idx = entry_index(entries, &part)
        .ok_or_else(|| anyhow!("SELECTOR_NOT_FOUND: worksheet {part} was not found"))?;
    if entry_text(entries, idx)?.contains("<drawing ") {
        bail!("FEATURE_NOT_IMPLEMENTED: XLSX addChart currently supports one drawing per sheet");
    }
    let chart_id = next_chart_id(entries);
    entries.push((
        format!("xl/drawings/drawing{chart_id}.xml"),
        drawing_xml(chart_id).into_bytes(),
    ));
    entries.push((
        format!("xl/drawings/_rels/drawing{chart_id}.xml.rels"),
        drawing_rels_xml(chart_id).into_bytes(),
    ));
    entries.push((
        format!("xl/charts/chart{chart_id}.xml"),
        chart_xml(op, &format!("Chart {chart_id}")).into_bytes(),
    ));
    upsert_content_type(
        entries,
        &format!("/xl/drawings/drawing{chart_id}.xml"),
        "application/vnd.openxmlformats-officedocument.drawing+xml",
    )?;
    upsert_content_type(
        entries,
        &format!("/xl/charts/chart{chart_id}.xml"),
        "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
    )?;
    let rel_id = add_sheet_relationship(
        entries,
        &part,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
        &format!("../drawings/drawing{chart_id}.xml"),
    )?;
    let sheet = ensure_worksheet_r_namespace(&entry_text(entries, idx)?);
    let next = if sheet.contains("<drawing ") {
        sheet
    } else {
        insert_before(
            &sheet,
            "</worksheet>",
            &format!(r#"<drawing r:id="rId{rel_id}"/>"#),
        )?
    };
    entries[idx].1 = next.into_bytes();
    Ok(true)
}

fn set_cells_in_sheet_xml(xml: &str, cells: &BTreeMap<String, XlsxCell>) -> String {
    let mut existing = BTreeMap::new();
    let cell_re = Regex::new(r#"(?s)<c\s+[^>]*r="([A-Za-z]+[0-9]+)"[^>]*>.*?</c>"#).unwrap();
    for cap in cell_re.captures_iter(xml) {
        if let (Some(full), Some(cell)) = (cap.get(0), cap.get(1)) {
            existing.insert(
                cell.as_str().to_ascii_uppercase(),
                full.as_str().to_string(),
            );
        }
    }
    for (cell, value) in cells {
        existing.insert(cell.to_ascii_uppercase(), cell_xml(cell, value));
    }
    let mut by_row: BTreeMap<usize, Vec<(String, String)>> = BTreeMap::new();
    for (cell, xml) in existing {
        let (_, row) = split_cell_ref(&cell).unwrap_or((1, 1));
        by_row.entry(row).or_default().push((cell, xml));
    }
    let rows = by_row
        .into_iter()
        .map(|(row, mut row_cells)| {
            row_cells
                .sort_by_key(|(cell, _)| split_cell_ref(cell).map(|(col, _)| col).unwrap_or(1));
            let cells_xml = row_cells
                .into_iter()
                .map(|(_, xml)| xml)
                .collect::<String>();
            format!(r#"<row r="{row}">{cells_xml}</row>"#)
        })
        .collect::<String>();
    let next_sheet_data = format!("<sheetData>{rows}</sheetData>");
    let re = Regex::new(r"(?s)<sheetData>.*?</sheetData>|<sheetData\s*/>").unwrap();
    if re.is_match(xml) {
        re.replace(xml, next_sheet_data).to_string()
    } else {
        insert_before(xml, "</worksheet>", &next_sheet_data).unwrap_or_else(|_| xml.to_string())
    }
}

fn add_rows_to_cells(
    cells: &mut BTreeMap<String, XlsxCell>,
    start: &str,
    rows: &[Value],
) -> Result<()> {
    if let Ok((start_col, start_row)) = split_cell_ref(start) {
        for (r_idx, row) in rows.iter().enumerate() {
            if let Some(values) = row.as_array() {
                for (c_idx, value) in values.iter().enumerate() {
                    cells.insert(
                        cell_ref(start_col + c_idx, start_row + r_idx),
                        cell_from_value(value)?,
                    );
                }
            }
        }
    }
    Ok(())
}

fn cell_from_value(value: &Value) -> Result<XlsxCell> {
    if let Some(formula) = value.get("formula").and_then(Value::as_str) {
        validate_formula_safety(formula)?;
        Ok(XlsxCell::Formula(formula.to_string()))
    } else {
        Ok(XlsxCell::Value(value.clone()))
    }
}

fn cell_xml(cell: &str, value: &XlsxCell) -> String {
    match value {
        XlsxCell::Formula(formula) => {
            format!(r#"<c r="{}"><f>{}</f></c>"#, cell, xml_escape(formula))
        }
        XlsxCell::Value(value) => {
            if let Some(n) = value.as_f64() {
                format!(r#"<c r="{cell}"><v>{n}</v></c>"#)
            } else if let Some(b) = value.as_bool() {
                format!(
                    r#"<c r="{cell}" t="b"><v>{}</v></c>"#,
                    if b { 1 } else { 0 }
                )
            } else {
                format!(
                    r#"<c r="{cell}" t="str"><v>{}</v></c>"#,
                    xml_escape(&cell_value_text(value))
                )
            }
        }
    }
}

fn cell_value_text(value: &Value) -> String {
    match value {
        Value::String(s) => s.clone(),
        Value::Null => String::new(),
        other => other.to_string(),
    }
}

fn range_cells(start: &str, rows: &[Value]) -> Result<BTreeMap<String, XlsxCell>> {
    let (start_col, start_row) = split_cell_ref(start)?;
    let mut out = BTreeMap::new();
    for (r_idx, row) in rows.iter().enumerate() {
        let values = row
            .as_array()
            .ok_or_else(|| anyhow!("SCHEMA_INVALID: xlsx.setRange values must be 2D array"))?;
        for (c_idx, value) in values.iter().enumerate() {
            out.insert(
                cell_ref(start_col + c_idx, start_row + r_idx),
                cell_from_value(value)?,
            );
        }
    }
    Ok(out)
}

fn split_cell_ref(cell: &str) -> Result<(usize, usize)> {
    let re = Regex::new(r"^([A-Za-z]+)([0-9]+)$").unwrap();
    let cap = re
        .captures(cell)
        .ok_or_else(|| anyhow!("SCHEMA_INVALID: invalid XLSX cell reference {cell}"))?;
    let col = cap[1].chars().fold(0usize, |acc, ch| {
        acc * 26 + (ch.to_ascii_uppercase() as u8 - b'A' + 1) as usize
    });
    let row = cap[2].parse::<usize>()?;
    Ok((col, row))
}

fn cell_ref(mut col: usize, row: usize) -> String {
    let mut letters = Vec::new();
    while col > 0 {
        col -= 1;
        letters.push((b'A' + (col % 26) as u8) as char);
        col /= 26;
    }
    letters.iter().rev().collect::<String>() + &row.to_string()
}

fn validations_xml(validations: &[Value]) -> String {
    if validations.is_empty() {
        return String::new();
    }
    let items = validations
        .iter()
        .filter_map(data_validation_xml)
        .collect::<String>();
    if items.is_empty() {
        String::new()
    } else {
        format!(
            r#"<dataValidations count="{}">{items}</dataValidations>"#,
            validations.len()
        )
    }
}

fn data_validation_xml(validation: &Value) -> Option<String> {
    let range = validation
        .get("range")
        .or_else(|| validation.get("sqref"))
        .and_then(Value::as_str)?;
    let kind = validation
        .get("type")
        .and_then(Value::as_str)
        .unwrap_or("list");
    let formula1 = validation
        .get("formula1")
        .or_else(|| validation.get("source"))
        .and_then(Value::as_str)
        .unwrap_or("");
    Some(format!(
        r#"<dataValidation type="{}" allowBlank="1" sqref="{}"><formula1>{}</formula1></dataValidation>"#,
        xml_escape(kind),
        xml_escape(range),
        xml_escape(formula1)
    ))
}

fn chart_data(spec: &Value) -> Result<(Vec<String>, Vec<f64>)> {
    if let (Some(labels), Some(values)) = (
        spec.get("labels").and_then(Value::as_array),
        spec.get("values").and_then(Value::as_array),
    ) {
        let labels = labels
            .iter()
            .map(|v| {
                v.as_str()
                    .map(str::to_string)
                    .unwrap_or_else(|| v.to_string())
            })
            .collect::<Vec<_>>();
        let values = values.iter().filter_map(Value::as_f64).collect::<Vec<_>>();
        if labels.len() == values.len() && !labels.is_empty() {
            return Ok((labels, values));
        }
    }
    if let Some(data) = spec.get("data") {
        if let Some(rows) = data.as_array() {
            if rows.iter().all(Value::is_array) {
                let mut labels = Vec::new();
                let mut values = Vec::new();
                for row in rows {
                    let row = row.as_array().unwrap();
                    if row.len() >= 2 {
                        labels.push(
                            row[0]
                                .as_str()
                                .map(str::to_string)
                                .unwrap_or_else(|| row[0].to_string()),
                        );
                        if let Some(value) = row[1].as_f64() {
                            values.push(value);
                        }
                    }
                }
                if labels.len() == values.len() && !labels.is_empty() {
                    return Ok((labels, values));
                }
            }
            let labels = rows
                .iter()
                .filter_map(|r| {
                    r.get("label")
                        .or_else(|| r.get("name"))
                        .and_then(Value::as_str)
                })
                .map(str::to_string)
                .collect::<Vec<_>>();
            let values = rows
                .iter()
                .filter_map(|r| {
                    r.get("value")
                        .or_else(|| r.get("y"))
                        .and_then(Value::as_f64)
                })
                .collect::<Vec<_>>();
            if labels.len() == values.len() && !labels.is_empty() {
                return Ok((labels, values));
            }
        }
        if let (Some(labels), Some(values)) = (
            data.get("labels").and_then(Value::as_array),
            data.get("values").and_then(Value::as_array),
        ) {
            let labels = labels
                .iter()
                .map(|v| {
                    v.as_str()
                        .map(str::to_string)
                        .unwrap_or_else(|| v.to_string())
                })
                .collect::<Vec<_>>();
            let values = values.iter().filter_map(Value::as_f64).collect::<Vec<_>>();
            if labels.len() == values.len() && !labels.is_empty() {
                return Ok((labels, values));
            }
        }
    }
    bail!("SCHEMA_INVALID: chart data requires labels/values or data rows");
}

fn worksheet_part_for_op(entries: &[(String, Vec<u8>)], op: &Value) -> Result<String> {
    if let Some(source_path) = op.pointer("/selector/sourcePath").and_then(Value::as_str) {
        return Ok(source_path.to_string());
    }
    if let Some(sheet) = op
        .get("sheet")
        .or_else(|| op.pointer("/selector/sheet"))
        .and_then(Value::as_str)
    {
        return sheet_part_by_name(entries, sheet);
    }
    if let Some(sheet_index) = op
        .get("sheet")
        .or_else(|| op.pointer("/selector/sheet"))
        .and_then(Value::as_u64)
    {
        return sheet_part_by_index(entries, sheet_index as usize);
    }
    let worksheets = entries
        .iter()
        .map(|(name, _)| name)
        .filter(|name| name.starts_with("xl/worksheets/sheet") && name.ends_with(".xml"))
        .cloned()
        .collect::<Vec<_>>();
    match worksheets.len() {
        0 => bail!("SELECTOR_NOT_FOUND: no worksheet parts found"),
        1 => Ok(worksheets[0].clone()),
        _ => bail!("SELECTOR_AMBIGUOUS: provide selector.sheet or selector.sourcePath"),
    }
}

fn worksheet_cell_for_op(entries: &[(String, Vec<u8>)], op: &Value) -> Result<(String, String)> {
    if let Some(cell) = op
        .get("cell")
        .or_else(|| op.pointer("/selector/cell"))
        .and_then(Value::as_str)
    {
        return Ok((
            worksheet_part_for_op(entries, op)?,
            cell.to_ascii_uppercase(),
        ));
    }
    if let Some(stable_id) = op
        .pointer("/selector/stableObjectId")
        .and_then(Value::as_str)
    {
        for (part, data) in entries
            .iter()
            .filter(|(name, _)| name.starts_with("xl/worksheets/") && name.ends_with(".xml"))
        {
            let xml = String::from_utf8_lossy(data);
            for cell in worksheet_cell_refs(&xml) {
                if stable_id_xlsx(part, &cell) == stable_id {
                    return Ok((part.clone(), cell));
                }
            }
        }
        bail!("SELECTOR_NOT_FOUND: XLSX stableObjectId did not match any worksheet cell");
    }
    bail!("SCHEMA_INVALID: xlsx cell edit requires cell, selector.cell, or selector.stableObjectId")
}

fn worksheet_cell_refs(xml: &str) -> Vec<String> {
    let cell_re = Regex::new(r#"<c\b[^>]*\br="([A-Za-z]+[0-9]+)""#).unwrap();
    cell_re
        .captures_iter(xml)
        .filter_map(|cap| cap.get(1))
        .map(|m| m.as_str().to_ascii_uppercase())
        .collect()
}

fn stable_id_xlsx(part: &str, cell: &str) -> String {
    let mut hash = Sha256::new();
    hash.update(b"xlsx");
    hash.update(part.as_bytes());
    hash.update(b":");
    hash.update(cell.to_ascii_uppercase().as_bytes());
    hash.update(b"0");
    format!("xlsx:{}", &hex::encode(hash.finalize())[..16])
}

fn sheet_part_by_name(entries: &[(String, Vec<u8>)], sheet: &str) -> Result<String> {
    let workbook = entries
        .iter()
        .find(|(name, _)| name == "xl/workbook.xml")
        .map(|(_, data)| String::from_utf8_lossy(data).to_string())
        .ok_or_else(|| anyhow!("SELECTOR_NOT_FOUND: xl/workbook.xml"))?;
    let re = Regex::new(&format!(
        r#"<sheet\b[^>]*\bname="{}"[^>]*\br:id="([^"]+)""#,
        regex::escape(&xml_escape(sheet))
    ))?;
    let rid = re
        .captures(&workbook)
        .and_then(|cap| cap.get(1))
        .map(|m| m.as_str().to_string())
        .ok_or_else(|| anyhow!("SELECTOR_NOT_FOUND: sheet {sheet}"))?;
    let rels = entries
        .iter()
        .find(|(name, _)| name == "xl/_rels/workbook.xml.rels")
        .map(|(_, data)| String::from_utf8_lossy(data).to_string())
        .ok_or_else(|| anyhow!("SELECTOR_NOT_FOUND: xl/_rels/workbook.xml.rels"))?;
    let re = Regex::new(&format!(
        r#"<Relationship\b[^>]*\bId="{}"[^>]*\bTarget="([^"]+)""#,
        regex::escape(&rid)
    ))?;
    let target = re
        .captures(&rels)
        .and_then(|cap| cap.get(1))
        .map(|m| m.as_str())
        .ok_or_else(|| anyhow!("SELECTOR_NOT_FOUND: workbook relationship for {sheet}"))?;
    Ok(normalize_workbook_rel_target(target))
}

fn sheet_part_by_index(entries: &[(String, Vec<u8>)], sheet_index: usize) -> Result<String> {
    let one_based = sheet_index.max(1);
    let worksheets = entries
        .iter()
        .map(|(name, _)| name)
        .filter(|name| name.starts_with("xl/worksheets/sheet") && name.ends_with(".xml"))
        .cloned()
        .collect::<Vec<_>>();
    worksheets
        .get(one_based - 1)
        .cloned()
        .ok_or_else(|| anyhow!("SELECTOR_NOT_FOUND: worksheet index {one_based}"))
}

fn normalize_workbook_rel_target(target: &str) -> String {
    let trimmed = target.trim_start_matches('/');
    if trimmed.starts_with("xl/") {
        trimmed.to_string()
    } else {
        format!("xl/{trimmed}")
    }
}

fn add_sheet_relationship(
    entries: &mut Vec<(String, Vec<u8>)>,
    worksheet_part: &str,
    rel_type: &str,
    target: &str,
) -> Result<usize> {
    let rels_part = worksheet_rels_part(worksheet_part)?;
    let idx = if let Some(idx) = entry_index(entries, &rels_part) {
        idx
    } else {
        entries.push((
            rels_part.clone(),
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>"#.to_vec(),
        ));
        entries.len() - 1
    };
    let rels = entry_text(entries, idx)?;
    let rid = next_rid(&rels);
    let rel = format!(r#"<Relationship Id="rId{rid}" Type="{rel_type}" Target="{target}"/>"#);
    entries[idx].1 = insert_before(&rels, "</Relationships>", &rel)?.into_bytes();
    Ok(rid)
}

fn worksheet_rels_part(worksheet_part: &str) -> Result<String> {
    let file = worksheet_part
        .rsplit('/')
        .next()
        .ok_or_else(|| anyhow!("SCHEMA_INVALID: invalid worksheet part"))?;
    Ok(format!("xl/worksheets/_rels/{file}.rels"))
}

fn upsert_content_type(
    entries: &mut Vec<(String, Vec<u8>)>,
    part_name: &str,
    content_type: &str,
) -> Result<()> {
    let idx = entry_index(entries, "[Content_Types].xml")
        .ok_or_else(|| anyhow!("SELECTOR_NOT_FOUND: [Content_Types].xml"))?;
    let content = entry_text(entries, idx)?;
    if content.contains(&format!(r#"PartName="{part_name}""#)) {
        return Ok(());
    }
    let override_xml =
        format!(r#"<Override PartName="{part_name}" ContentType="{content_type}"/>"#);
    entries[idx].1 = insert_before(&content, "</Types>", &override_xml)?.into_bytes();
    Ok(())
}

fn ensure_worksheet_r_namespace(sheet: &str) -> String {
    if sheet.contains("xmlns:r=") {
        return sheet.to_string();
    }
    sheet.replacen(
        "<worksheet ",
        r#"<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" "#,
        1,
    )
}

fn update_table_parts_count(xml: &str) -> String {
    let count = Regex::new(r#"<tablePart\b"#)
        .unwrap()
        .find_iter(xml)
        .count();
    Regex::new(r#"<tableParts\b[^>]*count="[0-9]+""#)
        .unwrap()
        .replace(xml, format!(r#"<tableParts count="{count}""#))
        .to_string()
}

fn update_data_validations_count(xml: &str) -> String {
    let count = Regex::new(r#"<dataValidation\b"#)
        .unwrap()
        .find_iter(xml)
        .count();
    Regex::new(r#"<dataValidations\b[^>]*count="[0-9]+""#)
        .unwrap()
        .replace(xml, format!(r#"<dataValidations count="{count}""#))
        .to_string()
}

fn entry_index(entries: &[(String, Vec<u8>)], name: &str) -> Option<usize> {
    entries.iter().position(|(entry, _)| entry == name)
}

fn entry_text(entries: &[(String, Vec<u8>)], idx: usize) -> Result<String> {
    String::from_utf8(entries[idx].1.clone()).map_err(Into::into)
}

fn worksheet_count(entries: &[(String, Vec<u8>)]) -> usize {
    entries
        .iter()
        .filter(|(name, _)| name.starts_with("xl/worksheets/sheet") && name.ends_with(".xml"))
        .count()
}

fn next_table_id(entries: &[(String, Vec<u8>)]) -> usize {
    next_part_id(entries, "xl/tables/table", ".xml")
}

fn next_chart_id(entries: &[(String, Vec<u8>)]) -> usize {
    next_part_id(entries, "xl/charts/chart", ".xml")
}

fn next_part_id(entries: &[(String, Vec<u8>)], prefix: &str, suffix: &str) -> usize {
    entries
        .iter()
        .filter_map(|(name, _)| {
            name.strip_prefix(prefix)
                .and_then(|rest| rest.strip_suffix(suffix))
                .and_then(|id| id.parse::<usize>().ok())
        })
        .max()
        .unwrap_or(0)
        + 1
}

fn next_rid(xml: &str) -> usize {
    Regex::new(r#"Id="rId([0-9]+)""#)
        .unwrap()
        .captures_iter(xml)
        .filter_map(|cap| cap.get(1).and_then(|m| m.as_str().parse::<usize>().ok()))
        .max()
        .unwrap_or(0)
        + 1
}

fn table_part_id(sheets: &[XlsxSheet], sheet_idx: usize, table_idx: usize) -> usize {
    sheets
        .iter()
        .take(sheet_idx)
        .map(|sheet| sheet.tables.len())
        .sum::<usize>()
        + table_idx
        + 1
}

fn chart_part_id(sheets: &[XlsxSheet], sheet_idx: usize, chart_idx: usize) -> usize {
    sheets
        .iter()
        .take(sheet_idx)
        .map(|sheet| sheet.charts.len())
        .sum::<usize>()
        + chart_idx
        + 1
}

fn drawing_rel_id(sheet: &XlsxSheet) -> usize {
    sheet.tables.len() + 1
}

fn table_rel_id(table_idx: usize) -> usize {
    table_idx + 1
}

fn table_column_count(range: &str) -> Result<usize> {
    let (_, end) = range
        .split_once(':')
        .ok_or_else(|| anyhow!("SCHEMA_INVALID: table range must be A1:B2 style"))?;
    let (start, _) = range.split_once(':').unwrap();
    let (start_col, _) = split_cell_ref(start)?;
    let (end_col, _) = split_cell_ref(end)?;
    Ok(end_col.saturating_sub(start_col) + 1)
}

fn valid_sheet_name(name: &str) -> String {
    let cleaned = name
        .chars()
        .map(|ch| {
            if matches!(ch, ':' | '\\' | '/' | '?' | '*' | '[' | ']') {
                '_'
            } else {
                ch
            }
        })
        .take(31)
        .collect::<String>();
    if cleaned.trim().is_empty() {
        "Sheet".into()
    } else {
        cleaned
    }
}

fn valid_table_name(name: &str) -> String {
    let cleaned = name
        .chars()
        .filter(|ch| ch.is_ascii_alphanumeric() || *ch == '_')
        .collect::<String>();
    if cleaned.is_empty() {
        "Table".into()
    } else {
        cleaned
    }
}

fn insert_before(haystack: &str, needle: &str, insert: &str) -> Result<String> {
    let idx = haystack
        .rfind(needle)
        .ok_or_else(|| anyhow!("SCHEMA_INVALID: expected XML marker {needle}"))?;
    let mut out = String::with_capacity(haystack.len() + insert.len());
    out.push_str(&haystack[..idx]);
    out.push_str(insert);
    out.push_str(&haystack[idx..]);
    Ok(out)
}

fn zip_entries_to_bytes(entries: Vec<(String, Vec<u8>)>) -> Result<Vec<u8>> {
    let mut out = Cursor::new(Vec::new());
    {
        let mut writer = ZipWriter::new(&mut out);
        let options =
            SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated);
        for (name, data) in entries {
            writer.start_file(name, options)?;
            writer.write_all(&data)?;
        }
        writer.finish()?;
    }
    Ok(out.into_inner())
}

fn xml_escape(text: &str) -> String {
    text.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
}
