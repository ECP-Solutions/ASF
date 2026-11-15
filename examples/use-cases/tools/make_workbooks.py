"""Writes the sample workbooks of the use cases from tools/data.py (needs openpyxl).

    python3 examples/use-cases/tools/make_workbooks.py

The workbooks hold data only: import the modules named in README.md, then run the macros. Their
results are not stored: run_lo.py checks what the scripts give.
"""
import os
from openpyxl import Workbook
from openpyxl.styles import Font

import data

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def sheet(wb, title, rows, text_columns=()):
    ws = wb.create_sheet(title)
    for r, row in enumerate(rows, start=1):
        for c, v in enumerate(row, start=1):
            cell = ws.cell(row=r, column=c, value=v)
            if r == 1:
                cell.font = Font(bold=True)
            elif c in text_columns:
                cell.number_format = "@"
    for c in range(1, (len(rows[0]) if rows else 1) + 1):
        width = max(len(str(row[c - 1])) for row in rows if c - 1 < len(row)) if rows else 10
        ws.column_dimensions[ws.cell(row=1, column=c).column_letter].width = max(10, width + 2)
    return ws


def formula(fn, args):
    def arg(a):
        if isinstance(a, bool):
            return "TRUE" if a else "FALSE"
        if isinstance(a, (int, float)):
            return str(a)
        return '"' + a.replace('"', '""') + '"'
    return "=%s(%s)" % (fn, ",".join(arg(a) for a in args))


def main():
    wb = Workbook()
    wb.remove(wb.active)
    sheet(wb, "Rules", data.RULES)
    sheet(wb, "Orders", data.ORDERS)
    sheet(wb, "Discounts", [["Run RulesHost.ApplyDiscountRules to fill this sheet"]])
    wb.save(os.path.join(BASE, "rules", "discounts.xlsx"))

    wb = Workbook()
    wb.remove(wb.active)
    sheet(wb, "Sales", data.sales(), text_columns=(3,))  # the month as text, not a date
    sheet(wb, "Summary", [["Run SalesHost.SummarizeSales to fill this sheet"]])
    wb.save(os.path.join(BASE, "pipeline", "sales.xlsx"))

    wb = Workbook()
    wb.remove(wb.active)
    rows = [["Formula", "Result", "Expected"]]
    for fn, args, expected in data.REGEX:
        f = formula(fn, args)
        rows.append([f[1:], f, expected if not isinstance(expected, bool) else ("TRUE" if expected else "FALSE")])
    ws = sheet(wb, "Examples", rows)
    ws.column_dimensions["B"].width = 22
    wb.save(os.path.join(BASE, "regex", "regex-functions.xlsx"))
    print("rules/discounts.xlsx, pipeline/sales.xlsx, regex/regex-functions.xlsx")


if __name__ == "__main__":
    main()
