#!/usr/bin/env python3
"""Compile the fictional cash-flow case and retain every source disagreement."""
from __future__ import annotations

import argparse
from decimal import Decimal
from html import escape
import json
from pathlib import Path
import sys

from build_office import build_artifacts, compile_report, ROOT


def check_expected(plan: dict, expected: dict) -> list[dict]:
    cells = {(section["id"], row["id"], column): cell["amount"]
             for section in plan["sections"] for row in section["rows"]
             for column, cell in row["cells"].items()}
    differences = []
    for section, rows in expected.items():
        if section == "description":
            continue
        for row, columns in rows.items():
            for column, amount in columns.items():
                actual = cells.pop((section, row, column), None)
                if actual is None or Decimal(actual) != Decimal(amount):
                    differences.append({"section": section, "row": row, "column": column,
                                        "expected": amount, "actual": actual})
    for (section, row, column), amount in cells.items():
        differences.append({"section": section, "row": row, "column": column,
                            "expected": None, "actual": amount})
    return differences


def review_html(plan: dict, differences: list[dict]) -> str:
    failed = [r for r in plan["source_reconciliations"] if not r["passed"]]
    tables = []
    for section in plan["sections"]:
        columns = section["columns"]
        headers = "<th>Description</th>" + "".join(f"<th>{escape(c['label'])}</th>" for c in columns)
        rows = []
        for row in section["rows"]:
            amounts = []
            for col in columns:
                cell = row["cells"][col["id"]]
                amount = Decimal(cell["amount"])
                display = f"({abs(amount):,.2f})" if amount < 0 else f"{amount:,.2f}"
                title = escape("Source facts: " + ", ".join(cell["fact_ids"]), quote=True)
                amounts.append(f'<td title="{title}">{display}</td>')
            rows.append(f"<tr><th>{escape(row['label'])}</th>{''.join(amounts)}</tr>")
        tables.append(f"<section><h2>{escape(section['title'])}</h2><table><thead><tr>{headers}</tr></thead><tbody>{''.join(rows)}</tbody></table></section>")
    status = f"{len(plan['checks'])} arithmetic controls passed · {len(failed)} source disagreements · {len(differences)} frozen expectation differences"
    exceptions = escape(json.dumps({"source_disagreements": failed, "expected_differences": differences}, indent=2))
    return """<!doctype html><html lang="en"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Harbor Tools — cash-flow reporting case</title><style>
body{margin:0;background:#f3f4f2;color:#182426;font:16px/1.5 system-ui,sans-serif}main{max-width:1180px;margin:auto;padding:48px 28px}h1{font-size:36px;line-height:1.15;margin:12px 0}h2{font-size:21px;margin:0 0 20px}.eyebrow{font-size:12px;font-weight:700;letter-spacing:.13em;color:#3d625b}p{max-width:860px}.status{padding:14px 18px;background:#e2ece7;font-weight:600}section{margin:32px 0;background:white;padding:26px;border:1px solid #d6dedb;overflow-x:auto}table{width:100%;border-collapse:collapse;font-size:14px}th,td{text-align:right;padding:10px 9px;border-bottom:1px solid #e2e7e4;white-space:nowrap;font-variant-numeric:tabular-nums}th:first-child{text-align:left;white-space:normal;min-width:270px}thead th{background:#edf1ee;font-size:12px}tbody th{font-weight:500}tr:has(td):nth-child(6),tr:last-child{font-weight:700}pre{white-space:pre-wrap;font-size:13px}a{color:#175d50}footer{font-size:13px;color:#465b55}@media(max-width:600px){main{padding:24px 14px}h1{font-size:29px}section{padding:18px 12px}th:first-child{min-width:220px}}</style>
<main><div class="eyebrow">FICTIONAL ACCOUNTING CASE · USD</div><h1>Harbor Tools consolidated cash flows</h1>
<p>Year ended June 30, 2025. A wholly owned parent and subsidiary, one eliminated internal loan, and one equipment purchase financed without cash. All workbooks and bank-control amounts were authored for this demonstration.</p>
<p class="status">""" + escape(status) + """</p><p><a href="report.xlsx">Numeric formula workbook</a> · <a href="report.docx">Report document</a> · <a href="report-plan.json">Exact amounts and source trace</a></p>
<p>Compiler predictions populate numeric formula caches. This view does not establish spreadsheet-engine recalculation or accounting approval. Restricted cash is collateral under a fictional lending covenant. Transfers into restricted cash are excluded from activity sections.</p>""" + "".join(tables) + f"<section><h2>Retained exceptions</h2><pre>{exceptions}</pre></section>" + """<footer>Equipment acquired through a $30,000 seller note appears only in noncash disclosures and debt/equipment rollforwards. Source hashes, coordinates, caption anchors and every formula dependency remain in the trace. This is a demonstration of cash-flow/support schedules, not a complete set of financial statements.</footer></main></html>"""


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--manifest", type=Path, default=ROOT / "examples/harbor/manifest.json")
    parser.add_argument("--expected", type=Path, default=ROOT / "examples/harbor/expected.json")
    parser.add_argument("--out", type=Path, default=ROOT / "output")
    args = parser.parse_args(argv)
    try:
        if args.out.exists() and any(args.out.iterdir()):
            raise ValueError("output directory is not empty; choose a fresh --out directory to preserve the previous generation")
        plan = compile_report(args.manifest, args.out, ROOT / "tools/report-compiler/target/release/acfr-report-compiler")
        differences = check_expected(plan, json.loads(args.expected.read_text()))
        record = build_artifacts(plan, args.out, title="Harbor Tools Inc. and subsidiary — fictional case",
            narrative="Restricted cash is collateral under a fictional lending covenant. Internal cash transfers are excluded from activity sections. Noncash equipment acquired through a seller note is disclosed separately. This report is illustrative and has not received accounting approval.")
        summary = {"schema": "cash-flow-reporting.review.v1", "manifest_sha256": plan["manifest_sha256"],
            "generation_sha256": record["generation_sha256"], "arithmetic_controls": len(plan["checks"]),
            "local_arithmetic_passed": plan["local_arithmetic_passed"],
            "source_reconciliations_passed": all(r["passed"] for r in plan["source_reconciliations"]),
            "source_reconciliations": plan["source_reconciliations"],
            "frozen_expectations_passed": not differences, "frozen_expectation_differences": differences,
            "formula_cache_origin": "same-generation exact-decimal compiler predictions; no spreadsheet-engine recalculation claimed",
            "accounting_approved": False}
        (args.out / "review.json").write_text(json.dumps(summary, indent=2) + "\n")
        (args.out / "index.html").write_text(review_html(plan, differences))
        print(json.dumps(summary, indent=2))
        return 0 if summary["source_reconciliations_passed"] and not differences else 1
    except (ValueError, OSError, KeyError) as exc:
        print(str(exc), file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
