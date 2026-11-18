"""Writes Pipeline.xlsx: the sample data, the settings of every stage, and one sheet per stage with
its ASF_CALL or ASF_RUN as an array formula, each reading the sheet of the stage before.

    python3 tools/make_workbook.py

Array formulas over fixed ranges work in Excel 2016 and later alike; rows added to Raw (up to row
RAW_ROWS) flow through every stage on recalculation. Needs openpyxl."""
import os, sys

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.formula import ArrayFormula

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import data  # noqa: E402

OUT = os.path.join(os.path.dirname(HERE), "Pipeline.xlsx")
RAW_ROWS = 200          # the rows of Raw the pipeline reads
BOLD = Font(bold=True)
HEAD = PatternFill("solid", fgColor="DDEBF7")
DATE = "yyyy-mm-dd"
MONEY = "#,##0.00"

# (sheet, cells of its array formula, formula, what it shows, number formats by column letter)
STAGES = [
    ("Clean", "A1:H200", '=ASF_CALL("clean",Raw!A1:H200,Settings!B2)',
     "Every text trimmed, the columns read as Settings!B2 says: dates and numbers from any format, codes in upper case", {"B": DATE, "G": MONEY}),
    ("Unique", "A1:H200", '=ASF_CALL("dedupe",Clean!A1:H200,Settings!B3,Settings!B4)',
     "One row per order: the last one of each, as a correction comes after the first", {"B": DATE, "G": MONEY}),
    ("Issues", "A1:E200", '=ASF_CALL("validate",Unique!A1:H200,Rules!A1:C50)',
     "What breaks the rules of the Rules sheet: one line per problem", {}),
    ("Valid", "A1:H200", '=ASF_CALL("validate",Unique!A1:H200,Rules!A1:C50,"valid")',
     "The rows with no problem: what the next stages use", {"B": DATE, "G": MONEY}),
    ("Enriched", "A1:L200", '=ASF_CALL("lookup",Valid!A1:H200,Customers!A1:D100,Settings!B5,Settings!B6)',
     "The customer's name, region and segment added; Matched is FALSE for an unknown customer", {"B": DATE, "G": MONEY}),
    ("Sales", "A1:M200", '=ASF_RUN(Scripts!B2,Enriched!A1:L200)',
     "Amount = Qty * Price, by the script in Scripts!B2 (ASF_RUN: a script in a cell)", {"B": DATE, "G": MONEY, "M": MONEY}),
    ("ByRegion", "A1:G30", '=ASF_CALL("group",Sales!A1:M200,Settings!B7,Settings!B8)',
     "One row per region: total, orders, average, customers, first and last date", {"B": MONEY, "D": MONEY, "F": DATE, "G": DATE}),
    ("Pivot", "A1:H20", '=ASF_CALL("pivot",Sales!A1:M200,Settings!B9,Settings!B10,Settings!B11)',
     "Amount by region and product, with totals", {c: MONEY for c in "BCDEFGH"}),
    ("Profile", "A1:J30", '=ASF_CALL("profile",Raw!A1:H200)',
     "What the raw data holds, column by column: the faults the Clean stage has to fix", {}),
]

SETTINGS = [
    ("Clean: how to read each column", data.CLEAN_SPEC),
    ("Dedupe: the key columns", "Order"),
    ("Dedupe: which row stays (first or last)", "last"),
    ("Lookup: the key (Column=ReferenceColumn)", "Customer=CustomerID"),
    ("Lookup: the columns to add", "Name, Region, Segment"),
    ("Group: by", "Region"),
    ("Group: the measures", "sum:Amount, count:Order, avg:Amount, distinct:Customer, min:Date, max:Date"),
    ("Pivot: rows", "Region"),
    ("Pivot: columns", "Product"),
    ("Pivot: the measure", "sum:Amount"),
]

START = [
    "ASF data pipeline",
    "",
    "Raw orders, as several sources export them, cleaned, checked, de-duplicated, enriched and summarized by",
    "worksheet formulas: every stage is one ASF_CALL (or ASF_RUN) reading the sheet of the stage before.",
    "Change Raw, Rules, Customers or Settings and every stage follows on recalculation.",
    "",
    "Before opening: load ASF.xlam, and copy the .vas files of the library folder that comes with this",
    "workbook into the ASF library folder (ASF tab > Library folder).",
    "",
    "Raw         the orders, as they come (up to row 200)",
    "Customers   the reference table the Enriched stage looks customers up in",
    "Rules       the validation rules: Column | Rule | Value",
    "Settings    the options of every stage",
    "Scripts     the script of the Sales stage, in a cell",
    "",
]


def table(ws, rows, start_row=1):
    for r, values in enumerate(rows, start=start_row):
        for c, v in enumerate(values, start=1):
            if v is not None:
                ws.cell(r, c, v)
    for c in range(1, len(rows[0]) + 1):
        ws.cell(start_row, c).font = BOLD
        ws.cell(start_row, c).fill = HEAD


def main():
    wb = Workbook()
    start = wb.active
    start.title = "Start"
    for r, line in enumerate(START, start=1):
        start.cell(r, 1, line)
    start["A1"].font = Font(bold=True, size=14)
    for i, (name, _, _, what, _) in enumerate(STAGES, start=1):
        start.cell(len(START) + i, 1, f"{i}. {name:<10} {what}")
    start.column_dimensions["A"].width = 110

    raw = wb.create_sheet("Raw")
    table(raw, data.RAW)
    customers = wb.create_sheet("Customers")
    table(customers, data.CUSTOMERS)
    rules = wb.create_sheet("Rules")
    table(rules, data.RULES)
    settings = wb.create_sheet("Settings")
    table(settings, [["Setting", "Value"]] + [list(s) for s in SETTINGS])
    scripts = wb.create_sheet("Scripts")
    table(scripts, [["Stage", "Script"], ["Sales", data.AMOUNT_SCRIPT]])
    scripts["B2"].alignment = Alignment(wrap_text=True, vertical="top")
    for ws, widths in ((raw, [10, 12, 10, 22, 12, 7, 11, 9]), (customers, [12, 12, 9, 11]), (rules, [10, 10, 34]),
                       (settings, [42, 90]), (scripts, [10, 100])):
        for i, w in enumerate(widths):
            ws.column_dimensions[chr(65 + i)].width = w

    for name, cells, formula, what, formats in STAGES:
        ws = wb.create_sheet(name)
        ws[cells.split(":")[0]] = ArrayFormula(cells, formula)
        first, last = cells.split(":")
        rows = int(last[1:])
        for col, fmt in formats.items():
            for r in range(2, rows + 1):
                ws[f"{col}{r}"].number_format = fmt
        for c in range(1, ord(last[0]) - 64 + 1):
            ws.cell(1, c).font = BOLD
            ws.column_dimensions[chr(64 + c)].width = 14
    wb.calculation.fullCalcOnLoad = True
    wb.save(OUT)
    print("wrote", OUT)


if __name__ == "__main__":
    main()
