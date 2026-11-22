"""Macros that work on the workbook through Excel's object model (xl_app, with pywin32). They run
from the ASF tab (Python macros), from VBA (Application.Run "name"), or from a button."""
import pandas
from asfpy import xl_app, xl_macro


@xl_macro(name='SummarizeSelection')
def summarize_selection():
    """Writes the count, mean, min and max of each numeric column of the selection beside it."""
    app = xl_app()
    sel = app.Selection
    values = sel.Value
    if not isinstance(values, tuple) or len(values) < 2:
        return 'Select a table with its header row first.'
    frame = pandas.DataFrame(list(values[1:]), columns=[str(h) for h in values[0]])
    stats = frame.apply(pandas.to_numeric, errors='coerce').agg(['count', 'mean', 'min', 'max']).dropna(axis=1, how='all')
    rows = [['', *stats.columns]] + [[name, *[round(float(x), 4) for x in stats.loc[name]]] for name in stats.index]
    target = sel.Offset(0, sel.Columns.Count + 1).Resize(len(rows), len(rows[0]))
    target.Value = rows
    target.Rows(1).Font.Bold = True
    return 'Summary written to %s' % target.Address


@xl_macro(name='ImportCsv')
def import_csv(path):
    """Reads a CSV file with pandas into a new sheet, its header in bold, the columns fitted."""
    frame = pandas.read_csv(path)
    app = xl_app()
    sheet = app.ActiveWorkbook.Worksheets.Add()
    rows = [list(frame.columns)] + frame.astype(object).where(frame.notna(), None).values.tolist()
    target = sheet.Range('A1').Resize(len(rows), len(rows[0]))
    target.Value = rows
    sheet.Rows(1).Font.Bold = True
    sheet.UsedRange.Columns.AutoFit()
    return '%d rows into %s' % (len(rows) - 1, sheet.Name)


@xl_macro(name='FitAllColumns')
def fit_all_columns():
    """Fits the columns of every sheet of the active workbook to their contents."""
    book = xl_app().ActiveWorkbook
    for sheet in book.Worksheets:
        sheet.UsedRange.Columns.AutoFit()
    return 'Columns fitted on %d sheets' % book.Worksheets.Count
