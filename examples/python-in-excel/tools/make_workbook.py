"""Writes PythonExamples.xlsx: the sample data, and one sheet per theme with the example functions
in formulas. Each array formula covers exactly the cells of its result, computed here through the
add-in's boot.py with the same data, so the workbook works in Excel 2016 (array formulas) and in
Microsoft 365 alike.

    python3 examples/python-in-excel/tools/make_workbook.py

Needs openpyxl, numpy and pandas."""
import datetime
import os
import sys

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.formula import ArrayFormula

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import check  # noqa: E402  (its boot.py loader)
import data  # noqa: E402

OUT = os.path.join(os.path.dirname(HERE), 'PythonExamples.xlsx')
BOLD = Font(bold=True)
TITLE = Font(bold=True, size=13)
NOTE = Font(italic=True, color='595959')
HEAD = PatternFill('solid', fgColor='DDEBF7')
INPUT = PatternFill('solid', fgColor='FFF2CC')

boot, names = check.load_boot()
SALES = [[check.serial(c) if isinstance(c, datetime.date) else c for c in row] for row in data.sales()]
MONTHLY = data.monthly()
REGIONS = data.regions()
TYPED, CLEAN = data.customers()


def size(name, *args):
    """The rows and columns of a function's result for the sample data."""
    r = check.cells(boot, boot.xl_call(names.index(name), list(args)))
    if isinstance(r, str) and r.startswith('FAIL'):
        raise RuntimeError('%s: %s' % (name, r))
    return (len(r), max(len(row) for row in r)) if isinstance(r, list) else (1, 1)


def col(rows, i):
    return [[r[i]] for r in rows[1:]]


def block(ws, row, column, label, formula, shape, note=''):
    """A label, the formula as an array formula over its result's cells (or a plain one for a
    single cell), and a note; returns the row after the block."""
    ws.cell(row, column, label).font = BOLD
    if note:
        ws.cell(row + 1, column, note).font = NOTE
    top = row + (2 if note else 1)
    rows, cols = shape
    ref = '%s%d' % (get_column_letter(column), top)
    if rows == 1 and cols == 1:
        ws[ref] = formula
    else:
        last = '%s%d' % (get_column_letter(column + cols - 1), top + rows - 1)
        ws[ref] = ArrayFormula('%s:%s' % (ref, last), formula)
    return top + rows + 1


def table(ws, row, column, rows, date_cols=()):
    for i, r in enumerate(rows):
        for j, v in enumerate(r):
            c = ws.cell(row + i, column + j, datetime.date(1899, 12, 30) + datetime.timedelta(days=v)
                        if j in date_cols and i and isinstance(v, float) else v)
            if i == 0:
                c.font, c.fill = BOLD, HEAD
            elif j in date_cols:
                c.number_format = 'yyyy-mm-dd'


def widths(ws, width=14):
    for i in range(1, 30):
        ws.column_dimensions[get_column_letter(i)].width = width


def main():
    wb = Workbook()
    start = wb.active
    start.title = 'Start'
    lines = [
        ('Python in Excel: examples', TITLE),
        ('These sheets call Python functions of the Python folder from formulas. To make them work:', None),
        ('1. Load the Python add-in (asfpy64.xll or asfpy32.xll): it carries Python with numpy, pandas, openpyxl and pywin32.', None),
        ('2. Copy the .py files of examples/python-in-excel/python into the Python folder (ASF tab: Python folder).', None),
        ('3. ASF tab: Reload Python, or restart Excel. Then Ctrl+Alt+F9 to calculate everything.', None),
        ('4. For the Scripts sheet: ASF.xlam, py_summary.vas in the library folder, and PythonAccess at call.', None),
        ('', None),
        ('Data: sample orders, a monthly series, a table of regions and names as typed.', None),
        ('Stats: numpy (z-scores, outliers, trend, seasonal forecast, moving average, histogram).', None),
        ('Tables: pandas (describe, group totals, pivot, query, join, top rows, correlation, a table kept as a handle).', None),
        ('Finance: a loan schedule, growth rate, and a Monte Carlo simulation that runs asynchronously.', None),
        ('Text: names tidied, accents removed, company names matched despite typos.', None),
        ('Scripts: an ASF script calling Python (ASF_CALL), and PY_CALL.', None),
        ('Macros (ASF tab: Python macros): SummarizeSelection, ImportCsv, FitAllColumns.', None),
        ('Python in a cell turns #NAME? until the add-in is loaded and the modules are in the Python folder.', NOTE),
    ]
    for i, (text, font) in enumerate(lines, 1):
        c = start.cell(i, 1, text)
        if font:
            c.font = font
    start.column_dimensions['A'].width = 110

    # --- Data
    ws = wb.create_sheet('Data')
    table(ws, 1, 1, SALES, date_cols=(1,))
    table(ws, 1, 9, MONTHLY)
    table(ws, 1, 12, REGIONS)
    table(ws, 1, 16, [['Typed']] + [[t] for t in TYPED])
    table(ws, 1, 18, [['Clean list']] + [[c] for c in CLEAN])
    widths(ws, 13)
    for c, w in (('A', 10), ('P', 20), ('R', 22)):
        ws.column_dimensions[c].width = w
    amount, qty, visitors = 'Data!G2:G61', 'Data!E2:E61', 'Data!J2:J25'
    sales, regions = 'Data!A1:G61', 'Data!L1:N5'

    # --- Stats: numpy
    ws = wb.create_sheet('Stats')
    block(ws, 1, 1, 'Z-scores of Amount', '=PY.ZSCORES(%s)' % amount, size('PY.ZSCORES', col(SALES, 6)), 'one per order')
    block(ws, 1, 2, 'Outlier?', '=PY.OUTLIERS(%s)' % qty, size('PY.OUTLIERS', col(SALES, 4)), "Qty, Tukey's fences")
    r = block(ws, 1, 4, 'Trend of Visitors', '=PY.TREND(%s)' % visitors, size('PY.TREND', col(MONTHLY, 1)), 'least squares')
    block(ws, r, 4, 'Next 12 months', '=PY.FORECAST(%s,12,12)' % visitors, size('PY.FORECAST', col(MONTHLY, 1), 12, 12),
          'trend + yearly pattern')
    block(ws, 1, 8, 'Moving average (3)', '=PY.MOVAVG(%s,3)' % visitors, size('PY.MOVAVG', col(MONTHLY, 1), 3), 'of Visitors')
    block(ws, 1, 10, 'Histogram of Qty', '=PY.HISTOGRAM(%s,4)' % qty, size('PY.HISTOGRAM', col(SALES, 4), 4), '4 intervals')
    widths(ws)

    # --- Tables: pandas
    ws = wb.create_sheet('Tables')
    r = block(ws, 1, 1, 'Summary of every numeric column', '=PY.DESCRIBE(%s)' % sales, size('PY.DESCRIBE', SALES), 'pandas describe')
    r = block(ws, r, 1, 'Amount by region', '=PY.GROUP(%s,"Region","Amount")' % sales, size('PY.GROUP', SALES, 'Region', 'Amount'))
    r = block(ws, r, 1, 'Pivot: Amount by region and product', '=PY.PIVOT(%s,"Region","Product","Amount")' % sales,
              size('PY.PIVOT', SALES, 'Region', 'Product', 'Amount'), 'with totals')
    r = block(ws, r, 1, 'Correlation of Qty, Price and Amount', '=PY.CORR(Data!E1:G61)',
              size('PY.CORR', [row[4:7] for row in SALES]))
    r = block(ws, r, 1, 'The 3 largest orders', '=PY.TOP(%s,"Amount",3)' % sales, size('PY.TOP', SALES, 'Amount', 3))
    cond = 'Region == "North" and Qty > 10'
    ws.cell(r, 1, 'Condition').font = BOLD
    ws.cell(r, 2, cond).fill = INPUT
    r = block(ws, r + 2, 1, 'Orders that meet the condition (edit B%d)' % r, '=PY.QUERY(%s,B%d)' % (sales, r),
              (61, 7), 'pandas query; cells past the result show #N/A')
    block(ws, 1, 9, 'The table kept in Python', '=PY.TABLE(%s)' % sales, (1, 1), 'a handle: other functions take this cell')
    block(ws, 4, 9, 'Qty by product, from the handle', '=PY.GROUP(I3,"Product","Qty")', size('PY.GROUP', SALES, 'Product', 'Qty'))
    block(ws, 10, 9, 'Orders with their region manager', '=PY.JOIN(%s,%s,"Region")' % (sales, regions),
          size('PY.JOIN', SALES, REGIONS, 'Region'), 'pandas merge')
    widths(ws)

    # --- Finance
    ws = wb.create_sheet('Finance')
    for i, (k, v, fmt) in enumerate((('Principal', 25000, '#,##0'), ('Annual rate', 0.065, '0.00%'), ('Years', 3, '0')), 1):
        ws.cell(i, 1, k).font = BOLD
        c = ws.cell(i, 2, v)
        c.fill, c.number_format = INPUT, fmt
    block(ws, 5, 1, 'Loan schedule', '=PY.LOAN(B1,B2,B3)', size('PY.LOAN', 25000, 0.065, 3), 'monthly payments')
    ws.cell(1, 8, 'Growth from 100 to 180 in 5 years').font = BOLD
    ws['H2'] = '=PY.CAGR(100,180,5)'
    ws['H2'].number_format = '0.00%'
    for i, (k, v, fmt) in enumerate((('Start', 10000, '#,##0'), ('Mean return', 0.06, '0.00%'),
                                     ('Volatility', 0.15, '0.00%'), ('Years', 10, '0'), ('Runs', 20000, '#,##0')), 4):
        ws.cell(i, 8, k).font = BOLD
        c = ws.cell(i, 9, v)
        c.fill, c.number_format = INPUT, fmt
    block(ws, 10, 8, 'Value after the years: percentiles', '=PY.MONTECARLO(I4,I5,I6,I7,I8)',
          size('PY.MONTECARLO', 10000, 0.06, 0.15, 10, 20000), 'asynchronous: Excel goes on meanwhile')
    widths(ws)

    # --- Text
    ws = wb.create_sheet('Text')
    for j, h in enumerate(('Typed', 'PY.PROPERNAME', 'PY.NOACCENTS', 'PY.FUZZY (clean list)', 'PY.SIMILARITY'), 1):
        c = ws.cell(1, j, h)
        c.font, c.fill = BOLD, HEAD
    for i, t in enumerate(TYPED, 2):
        ws.cell(i, 1, t)
        ws.cell(i, 2, '=PY.PROPERNAME(A%d)' % i)
        ws.cell(i, 3, '=PY.NOACCENTS(B%d)' % i)
        ws.cell(i, 4, '=PY.FUZZY(A%d,Data!R2:R6,0.5)' % i)
        ws.cell(i, 5, '=IF(D%d="","",PY.SIMILARITY(A%d,D%d))' % (i, i, i))
    widths(ws, 24)

    # --- Scripts: ASF and PY_CALL
    ws = wb.create_sheet('Scripts')
    block(ws, 1, 1, 'Summary of Amount, by an ASF script calling Python', '=ASF_CALL("py_summary",%s)' % amount, (6, 2),
          'library/py_summary.vas; needs ASF.xlam and PythonAccess at call')
    ws.cell(11, 1, 'Orders above the mean Amount, by PY_CALL').font = BOLD
    ws['A12'] = '=PY_CALL("xl_stats.above_mean",%s)' % amount
    widths(ws, 22)

    for sheet in wb.worksheets:
        for row in sheet.iter_rows():
            for c in row:
                if c.alignment is None or not c.alignment.wrap_text:
                    c.alignment = Alignment(vertical='top')
    wb.save(OUT)
    print(OUT)


if __name__ == '__main__':
    main()
