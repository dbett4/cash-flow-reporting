use calamine::{Data, Reader, open_workbook_auto_from_rs};
use rust_decimal::Decimal;
use serde::{Deserialize, Serialize};
use sha2::{Digest, Sha256};
use std::collections::{BTreeMap, BTreeSet};
use std::error::Error;
use std::fs;
use std::io::Cursor;
use std::path::Path;
use std::str::FromStr;

type Result<T> = std::result::Result<T, Box<dyn Error>>;

#[derive(Deserialize)]
#[serde(deny_unknown_fields)]
struct Source {
    id: String,
    path: String,
    sha256: String,
    sheet: String,
    period_cell: String,
}

#[derive(Deserialize)]
#[serde(deny_unknown_fields)]
struct FactSpec {
    id: String,
    source: String,
    cell: String,
    #[serde(default)]
    anchors: BTreeMap<String, String>,
}

#[derive(Deserialize, Clone)]
#[serde(tag = "kind", rename_all = "snake_case", deny_unknown_fields)]
enum Expression {
    Fact {
        id: String,
    },
    Sum {
        terms: Vec<Expression>,
    },
    Difference {
        left: Box<Expression>,
        right: Box<Expression>,
    },
    Reference {
        section: String,
        row: String,
        column: String,
    },
    // A declared empty source slot is an explicit non-applicable zero, never
    // a missing or unreadable financial input silently converted to zero.
    Zero {
        reason: String,
    },
}

#[derive(Deserialize)]
#[serde(deny_unknown_fields)]
struct Column {
    id: String,
    label: String,
}

#[derive(Deserialize)]
#[serde(deny_unknown_fields)]
struct Row {
    id: String,
    label: String,
    #[serde(default)]
    decimals: u8,
    values: BTreeMap<String, Expression>,
}

#[derive(Deserialize)]
#[serde(deny_unknown_fields)]
struct Section {
    id: String,
    title: String,
    #[serde(default)]
    columns: Vec<String>,
    rows: Vec<Row>,
}

#[derive(Deserialize)]
#[serde(deny_unknown_fields)]
struct Check {
    id: String,
    left: Expression,
    right: Expression,
    #[serde(default)]
    tolerance: Option<String>,
    #[serde(default)]
    residual_reason: Option<String>,
}

#[derive(Deserialize)]
#[serde(deny_unknown_fields)]
struct Manifest {
    schema: String,
    fiscal_year: u32,
    sources: Vec<Source>,
    facts: Vec<FactSpec>,
    columns: Vec<Column>,
    sections: Vec<Section>,
    checks: Vec<Check>,
    #[serde(default)]
    source_reconciliations: Vec<Check>,
}

#[derive(Serialize)]
struct Fact {
    id: String,
    source: String,
    source_sha256: String,
    sheet: String,
    cell: String,
    #[serde(with = "rust_decimal::serde::str")]
    amount: Decimal,
    cached_formula: Option<String>,
    anchors: BTreeMap<String, String>,
}

#[derive(Clone, Serialize)]
struct Cell {
    #[serde(with = "rust_decimal::serde::str")]
    amount: Decimal,
    formula: String,
    fact_ids: BTreeSet<String>,
    zero_reasons: BTreeSet<String>,
}

#[derive(Serialize)]
struct OutputRow {
    id: String,
    label: String,
    decimals: u8,
    cells: BTreeMap<String, Cell>,
}

#[derive(Serialize)]
struct OutputSection {
    id: String,
    title: String,
    columns: Vec<BTreeMap<&'static str, String>>,
    rows: Vec<OutputRow>,
}

fn section_columns<'a>(manifest: &'a Manifest, section: &Section) -> Result<Vec<&'a Column>> {
    let columns = if section.columns.is_empty() {
        manifest.columns.iter().collect::<Vec<_>>()
    } else {
        let mut seen = BTreeSet::new();
        section
            .columns
            .iter()
            .map(|id| {
                if !seen.insert(id) {
                    return Err("duplicate section column".into());
                }
                manifest
                    .columns
                    .iter()
                    .find(|c| &c.id == id)
                    .ok_or_else(|| "unknown section column".into())
            })
            .collect::<Result<Vec<_>>>()?
    };
    if columns.len() > 25 {
        return Err(
            "native report assembly supports at most 25 numeric columns per section".into(),
        );
    }
    Ok(columns)
}

#[derive(Serialize)]
struct CheckResult {
    id: String,
    #[serde(with = "rust_decimal::serde::str")]
    residual: Decimal,
    passed: bool,
    tolerance: String,
    residual_reason: Option<String>,
}

#[derive(Serialize)]
struct Output {
    schema: &'static str,
    fiscal_year: u32,
    manifest_sha256: String,
    facts: Vec<Fact>,
    columns: Vec<BTreeMap<&'static str, String>>,
    sections: Vec<OutputSection>,
    checks: Vec<CheckResult>,
    source_reconciliations: Vec<CheckResult>,
    local_arithmetic_passed: bool,
    native_verified: bool,
    client_ready: bool,
}

fn coordinate(cell: &str) -> Result<(u32, u32)> {
    let split = cell
        .find(|c: char| c.is_ascii_digit())
        .ok_or("cell requires an A1 coordinate")?;
    let (letters, digits) = cell.split_at(split);
    if letters.is_empty()
        || !letters.chars().all(|c| c.is_ascii_uppercase())
        || !digits.chars().all(|c| c.is_ascii_digit())
        || digits.starts_with('0')
    {
        return Err(format!("invalid A1 coordinate: {cell}").into());
    }
    let mut col = 0u32;
    for byte in letters.bytes() {
        col = col
            .checked_mul(26)
            .and_then(|n| n.checked_add((byte - b'A' + 1) as u32))
            .ok_or("column overflow")?;
    }
    let row = digits.parse::<u32>()?;
    if row > 1_048_576 || col > 16_384 {
        return Err("cell outside Excel limits".into());
    }
    Ok((row - 1, col - 1))
}

fn column_letter(mut n: usize) -> String {
    let mut out = Vec::new();
    n += 1;
    while n > 0 {
        n -= 1;
        out.push(b'A' + (n % 26) as u8);
        n /= 26;
    }
    out.reverse();
    String::from_utf8(out).unwrap()
}

fn numeric(value: &Data) -> Result<Decimal> {
    match value {
        Data::Int(v) => Ok(Decimal::from(*v)),
        Data::Float(v) if v.is_finite() => Ok(Decimal::from_str(&v.to_string())?),
        // Financial text, dates, errors and empty cells are not numeric facts.
        _ => Err("source fact is not a finite native number".into()),
    }
}

fn sha256(bytes: &[u8]) -> String {
    format!("{:x}", Sha256::digest(bytes))
}

fn validate_id(id: &str) -> Result<()> {
    if id.is_empty()
        || !id
            .chars()
            .all(|c| c.is_ascii_alphanumeric() || c == '_' || c == '-')
    {
        return Err(format!("invalid logical identifier: {id}").into());
    }
    Ok(())
}

fn load_facts(manifest: &Manifest, base: &Path) -> Result<Vec<Fact>> {
    let mut sources = BTreeMap::new();
    for source in &manifest.sources {
        validate_id(&source.id)?;
        if sources.contains_key(&source.id) {
            return Err("duplicate source id".into());
        }
        if source.sheet.trim().is_empty() {
            return Err("source requires an exact sheet name".into());
        }
        let path = base.join(&source.path);
        let bytes = fs::read(&path)?;
        if sha256(&bytes) != source.sha256 {
            return Err(format!("source hash mismatch: {}", source.id).into());
        }
        let mut workbook = open_workbook_auto_from_rs(Cursor::new(bytes))
            .map_err(|e| format!("cannot read source {}: {e}; encrypted or rights-managed inputs require an authorized readable source", source.id))?;
        let values = workbook.worksheet_range(&source.sheet)?;
        let formulas = workbook.worksheet_formula(&source.sheet)?;
        let period = values
            .get_value(coordinate(&source.period_cell)?)
            .ok_or("period cell is outside the source data")?;
        if period.to_string().trim() != manifest.fiscal_year.to_string() {
            return Err(format!("source period mismatch: {}", source.id).into());
        }
        sources.insert(source.id.clone(), (source, values, formulas));
    }
    let mut seen = BTreeSet::new();
    let mut facts = Vec::new();
    for spec in &manifest.facts {
        validate_id(&spec.id)?;
        if !seen.insert(spec.id.clone()) {
            return Err("duplicate fact id".into());
        }
        let (source, values, formulas) = sources
            .get(&spec.source)
            .ok_or_else(|| format!("unknown fact source: {}", spec.source))?;
        let coord = coordinate(&spec.cell)?;
        for (address, expected) in &spec.anchors {
            let label = values
                .get_value(coordinate(address)?)
                .ok_or("fact anchor is outside source data")?;
            let actual = match label {
                Data::String(text) => text.split_whitespace().collect::<Vec<_>>().join(" "),
                _ => {
                    return Err(
                        format!("fact anchor is not a caption: {}!{address}", source.id).into(),
                    );
                }
            };
            if expected.trim().is_empty()
                || actual != expected.split_whitespace().collect::<Vec<_>>().join(" ")
            {
                return Err(format!(
                    "fact caption changed: {}!{address} for {}",
                    source.id, spec.id
                )
                .into());
            }
        }
        let value = values
            .get_value(coord)
            .ok_or("fact cell is outside the source data")?;
        let amount = numeric(value).map_err(|e| format!("{}!{}: {e}", source.id, spec.cell))?;
        let formula = formulas.get_value(coord).filter(|v| !v.is_empty()).cloned();
        facts.push(Fact {
            id: spec.id.clone(),
            source: spec.source.clone(),
            source_sha256: source.sha256.clone(),
            sheet: source.sheet.clone(),
            cell: spec.cell.clone(),
            amount,
            cached_formula: formula,
            anchors: spec.anchors.clone(),
        });
    }
    Ok(facts)
}

struct Compiler<'a> {
    facts: &'a [Fact],
    manifest: &'a Manifest,
    cache: BTreeMap<(String, String, String), Cell>,
    active: BTreeSet<(String, String, String)>,
}

impl Compiler<'_> {
    fn reference(&mut self, section: &str, row: &str, column: &str) -> Result<Cell> {
        let key = (section.to_string(), row.to_string(), column.to_string());
        if let Some(cell) = self.cache.get(&key) {
            return Ok(cell.clone());
        }
        if !self.active.insert(key.clone()) {
            return Err(format!("cyclic report dependency: {key:?}").into());
        }
        let sec = self
            .manifest
            .sections
            .iter()
            .find(|s| s.id == section)
            .ok_or("unknown referenced section")?;
        let row_index = sec
            .rows
            .iter()
            .position(|r| r.id == row)
            .ok_or("unknown referenced row")?;
        let expr = sec.rows[row_index]
            .values
            .get(column)
            .ok_or("unknown referenced column")?
            .clone();
        let result = self.expression(&expr)?;
        self.active.remove(&key);
        self.cache.insert(key, result.clone());
        Ok(result)
    }

    fn expression(&mut self, expr: &Expression) -> Result<Cell> {
        match expr {
            Expression::Fact { id } => {
                let index = self
                    .facts
                    .iter()
                    .position(|f| &f.id == id)
                    .ok_or("unknown fact id")?;
                Ok(Cell {
                    amount: self.facts[index].amount,
                    formula: format!("'Source Facts'!F{}", index + 2),
                    fact_ids: BTreeSet::from([id.clone()]),
                    zero_reasons: BTreeSet::new(),
                })
            }
            Expression::Zero { reason } => {
                if reason.trim().is_empty() {
                    return Err("explicit zero requires a reason".into());
                }
                Ok(Cell {
                    amount: Decimal::ZERO,
                    formula: "0".into(),
                    fact_ids: BTreeSet::new(),
                    zero_reasons: BTreeSet::from([reason.clone()]),
                })
            }
            Expression::Sum { terms } => {
                if terms.is_empty() {
                    return Err("empty sum cannot stand in for missing evidence".into());
                }
                let cells = terms
                    .iter()
                    .map(|term| self.expression(term))
                    .collect::<Result<Vec<_>>>()?;
                let mut result = Cell {
                    amount: Decimal::ZERO,
                    formula: String::new(),
                    fact_ids: BTreeSet::new(),
                    zero_reasons: BTreeSet::new(),
                };
                for cell in &cells {
                    result.amount = result
                        .amount
                        .checked_add(cell.amount)
                        .ok_or("decimal overflow")?;
                    result.fact_ids.extend(cell.fact_ids.clone());
                    result.zero_reasons.extend(cell.zero_reasons.clone());
                }
                result.formula = format!(
                    "({})",
                    cells
                        .iter()
                        .map(|c| c.formula.as_str())
                        .collect::<Vec<_>>()
                        .join("+")
                );
                Ok(result)
            }
            Expression::Difference { left, right } => {
                let mut a = self.expression(left)?;
                let b = self.expression(right)?;
                a.amount = a.amount.checked_sub(b.amount).ok_or("decimal overflow")?;
                a.formula = format!("({}-{})", a.formula, b.formula);
                a.fact_ids.extend(b.fact_ids);
                a.zero_reasons.extend(b.zero_reasons);
                Ok(a)
            }
            Expression::Reference {
                section,
                row,
                column,
            } => {
                let mut result = self.reference(section, row, column)?;
                let sec = self
                    .manifest
                    .sections
                    .iter()
                    .find(|s| &s.id == section)
                    .ok_or("unknown section")?;
                let r = sec
                    .rows
                    .iter()
                    .position(|r| &r.id == row)
                    .ok_or("unknown row")?;
                let active_columns = section_columns(self.manifest, sec)?;
                let c = active_columns
                    .iter()
                    .position(|c| &c.id == column)
                    .ok_or("unknown column")?;
                result.formula = format!(
                    "'{}'!{}{}",
                    section.replace('\'', "''"),
                    column_letter(c + 1),
                    r + 5
                );
                Ok(result)
            }
        }
    }
}

fn compile(manifest: &Manifest, manifest_hash: String, facts: Vec<Fact>) -> Result<Output> {
    if manifest.schema != "acfr.report_manifest.v1" {
        return Err("unsupported report manifest schema".into());
    }
    if manifest.columns.is_empty() || manifest.sections.is_empty() || manifest.checks.is_empty() {
        return Err("report requires columns, sections and accounting checks".into());
    }
    let mut columns = BTreeSet::new();
    for col in &manifest.columns {
        validate_id(&col.id)?;
        if !columns.insert(&col.id) {
            return Err("duplicate column id".into());
        }
    }
    let mut sections = BTreeSet::new();
    for section in &manifest.sections {
        validate_id(&section.id)?;
        if section.id.len() > 31 || section.id == "Source Facts" || !sections.insert(&section.id) {
            return Err("invalid or duplicate section sheet id".into());
        }
        let mut rows = BTreeSet::new();
        let active_columns = section_columns(manifest, section)?;
        let expected_columns = active_columns
            .iter()
            .map(|c| &c.id)
            .collect::<BTreeSet<_>>();
        if section.rows.is_empty() {
            return Err("empty report section".into());
        }
        for row in &section.rows {
            validate_id(&row.id)?;
            if row.decimals > 6 {
                return Err("report row precision must be at most six decimals".into());
            }
            if !rows.insert(&row.id)
                || row.values.keys().collect::<BTreeSet<_>>() != expected_columns
            {
                return Err("duplicate row or incomplete column coverage".into());
            }
        }
    }
    let mut compiler = Compiler {
        facts: &facts,
        manifest,
        cache: BTreeMap::new(),
        active: BTreeSet::new(),
    };
    let mut output_sections = Vec::new();
    for section in &manifest.sections {
        let mut rows = Vec::new();
        for row in &section.rows {
            let mut cells = BTreeMap::new();
            for col in section_columns(manifest, section)? {
                cells.insert(
                    col.id.clone(),
                    compiler.reference(&section.id, &row.id, &col.id)?,
                );
            }
            rows.push(OutputRow {
                id: row.id.clone(),
                label: row.label.clone(),
                decimals: row.decimals,
                cells,
            });
        }
        output_sections.push(OutputSection {
            id: section.id.clone(),
            title: section.title.clone(),
            columns: section_columns(manifest, section)?
                .iter()
                .map(|c| BTreeMap::from([("id", c.id.clone()), ("label", c.label.clone())]))
                .collect(),
            rows,
        });
    }
    let mut check_ids = BTreeSet::new();
    let mut checks = Vec::new();
    let mut source_reconciliations = Vec::new();
    for (check, source_control) in manifest
        .checks
        .iter()
        .map(|c| (c, false))
        .chain(manifest.source_reconciliations.iter().map(|c| (c, true)))
    {
        validate_id(&check.id)?;
        if !check_ids.insert(&check.id) {
            return Err("duplicate check id".into());
        }
        let residual = compiler
            .expression(&check.left)?
            .amount
            .checked_sub(compiler.expression(&check.right)?.amount)
            .ok_or("decimal overflow")?;
        let tolerance = Decimal::from_str(check.tolerance.as_deref().unwrap_or("0"))?;
        if tolerance < Decimal::ZERO
            || tolerance > Decimal::from_str("0.000001")?
            || (tolerance != Decimal::ZERO
                && check
                    .residual_reason
                    .as_deref()
                    .unwrap_or("")
                    .trim()
                    .is_empty())
        {
            return Err(
                "only explicitly classified sub-cent source precision is admissible".into(),
            );
        }
        let result = CheckResult {
            id: check.id.clone(),
            residual,
            passed: residual.abs() <= tolerance,
            tolerance: tolerance.to_string(),
            residual_reason: check.residual_reason.clone(),
        };
        if source_control {
            source_reconciliations.push(result);
        } else {
            checks.push(result);
        }
    }
    let passed = checks.iter().all(|c| c.passed);
    Ok(Output {
        schema: "acfr.report_plan.v1",
        fiscal_year: manifest.fiscal_year,
        manifest_sha256: manifest_hash,
        facts,
        columns: manifest
            .columns
            .iter()
            .map(|c| BTreeMap::from([("id", c.id.clone()), ("label", c.label.clone())]))
            .collect(),
        sections: output_sections,
        checks,
        source_reconciliations,
        local_arithmetic_passed: passed,
        native_verified: false,
        client_ready: false,
    })
}

fn run() -> Result<()> {
    let args: Vec<String> = std::env::args().collect();
    if args.len() != 3 {
        return Err("usage: acfr-report-compiler MANIFEST.json OUTPUT.json".into());
    }
    let path = Path::new(&args[1]);
    let bytes = fs::read(path)?;
    let manifest: Manifest = serde_json::from_slice(&bytes)?;
    let facts = load_facts(&manifest, path.parent().unwrap_or(Path::new(".")))?;
    let output = compile(&manifest, sha256(&bytes), facts)?;
    fs::write(&args[2], serde_json::to_vec_pretty(&output)?)?;
    if !output.local_arithmetic_passed {
        return Err("accounting checks failed; report plan is not admitted".into());
    }
    eprintln!(
        "compiled {} sections, {} source facts, {} accounting checks; native verification required",
        output.sections.len(),
        output.facts.len(),
        output.checks.len()
    );
    Ok(())
}

fn main() {
    if let Err(err) = run() {
        eprintln!("{err}");
        std::process::exit(2);
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn exact_accounting_numbers_and_strict_source_types() {
        assert_eq!(
            numeric(&Data::Float(0.1)).unwrap() + numeric(&Data::Float(0.2)).unwrap(),
            Decimal::from_str("0.3").unwrap()
        );
        assert!(numeric(&Data::String("123".into())).is_err());
        assert!(numeric(&Data::Empty).is_err());
        assert!(numeric(&Data::Float(f64::NAN)).is_err());
    }
    #[test]
    fn addresses_are_absolute_and_bounded() {
        assert_eq!(coordinate("G54").unwrap(), (53, 6));
        assert_eq!(coordinate("XFD1048576").unwrap(), (1048575, 16383));
        for cell in ["G0", "G01", "g54", "XFE1", "A1048577", "A1:A2"] {
            assert!(coordinate(cell).is_err());
        }
    }
}
