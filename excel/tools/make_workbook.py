"""Writes excel/test/AddinTests.xlsx: the cases of cases.py as formulas, for Excel with the add-in.

    python3 excel/tools/make_workbook.py

Each case is one block of rows: its name (column A), whether it passes (B), the formula (from C),
the expected result (from L) and the cells its ranges read (from U). A result over several cells
is an array formula over them, as Excel 2016 and 2019 enter it (Ctrl+Shift+Enter); Microsoft 365
runs it the same way. B1 counts the failures. Open the workbook with the add-in loaded; it
recalculates on opening. The cases that need the library folder or a low step limit (lo_only) are
left to the LibreOffice checker. Needs openpyxl.
"""
import os, sys

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.formula import ArrayFormula

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from cases import CASES, Error, Range, shape  # noqa: E402

OUT = os.path.join(os.path.dirname(HERE), "test", "AddinTests.xlsx")
RESULT_COL, EXPECTED_COL, INPUT_COL = 3, 12, 21           # C, L, U
ERROR_TYPE = {"#NULL!": 1, "#DIV/0!": 2, "#VALUE!": 3, "#REF!": 4, "#NAME?": 5, "#NUM!": 6, "#N/A": 7}


def ref(row, col, rows=1, cols=1):
    a = f"{get_column_letter(col)}{row}"
    return a if rows == 1 and cols == 1 else f"{a}:{get_column_letter(col + cols - 1)}{row + rows - 1}"


def literal(v):
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, (int, float)):
        return repr(v)
    return '"' + str(v).replace('"', '""') + '"'


def cell_value(ws, row, col, v):
    if isinstance(v, Error):
        ws.cell(row, col, Error.FORMULAS[v.text])
    elif isinstance(v, str) and v in Error.CODES:
        # text that reads as an error value: openpyxl would store the error
        ws.cell(row, col, '="%s"' % v)
    elif v is not None and v != "":
        ws.cell(row, col, v)


def main():
    wb = Workbook()
    ws = wb.active
    ws.title = "Tests"
    ws["A1"] = "ASF add-in tests"
    ws["A1"].font = Font(bold=True)
    ws["A2"] = "Open with ASF.xlam loaded. B1: the result of all checks."
    for col, title in ((1, "Case"), (2, "Pass"), (RESULT_COL, "Result"), (EXPECTED_COL, "Expected"), (INPUT_COL, "Inputs")):
        ws.cell(4, col, title).font = Font(bold=True)
    row = 5
    first = row
    for case in CASES:
        if case.get("lo_only"):
            continue
        rows, cols = shape(case["expected"])
        args, height, input_col = [], rows, INPUT_COL
        for a in case["args"]:
            if isinstance(a, Range):
                for r, values in enumerate(a.rows):
                    for c, v in enumerate(values):
                        cell_value(ws, row + r, input_col + c, v)
                width = max(len(r) for r in a.rows)
                args.append(ref(row, input_col, len(a.rows), width))
                height = max(height, len(a.rows))
                input_col += width + 1
            else:
                args.append(literal(a))
        formula = f"={case['fn']}({', '.join(args)})"
        result = ref(row, RESULT_COL, rows, cols)
        expected = ref(row, EXPECTED_COL, rows, cols)
        ws.cell(row, 1, f"{case['fn']}: {case['name']}")
        if rows == 1 and cols == 1:
            ws.cell(row, RESULT_COL, formula)
        else:
            ws[ref(row, RESULT_COL)] = ArrayFormula(result, formula)
        exp = case["expected"]
        if isinstance(exp, Error):
            ws.cell(row, EXPECTED_COL, "error " + exp.text)
            ws.cell(row, 2, f"=IFERROR(ERROR.TYPE({result})={ERROR_TYPE[exp.text]},FALSE)")
        elif isinstance(exp, list):
            for r, values in enumerate(exp):
                for c, v in enumerate(values):
                    cell_value(ws, row + r, EXPECTED_COL + c, v)
            b = ref(row, 2)
            ws[b] = ArrayFormula(b, f"=AND(IFERROR(EXACT({result},{expected}),FALSE))")
        else:
            cell_value(ws, row, EXPECTED_COL, exp)
            ws.cell(row, 2, f"=IFERROR(EXACT({result},{expected}),FALSE)")
        row += height + 1
    ws["B1"] = (f'=IF(COUNTIF(B{first}:B{row},FALSE)=0,"All "&COUNTIF(B{first}:B{row},TRUE)&" checks pass",'
                f'COUNTIF(B{first}:B{row},FALSE)&" checks fail")')
    ws["B1"].font = Font(bold=True)
    ws.column_dimensions["A"].width = 52
    wb.calculation.fullCalcOnLoad = True
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    wb.save(OUT)
    print("wrote", OUT)


if __name__ == "__main__":
    main()
