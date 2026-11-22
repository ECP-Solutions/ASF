"""Run Python script: an overview of the active workbook, printed: each sheet with its used range,
how many cells hold formulas, and the functions its formulas call, most used first."""
import re

from asfpy import xl_app

app = xl_app()
book = app.ActiveWorkbook
print(book.Name)
python_calls = {}
for sheet in book.Worksheets:
    used = sheet.UsedRange
    formulas = used.Formula
    if not isinstance(formulas, tuple):
        formulas = ((formulas,),)
    cells = [f for row in formulas for f in row]
    with_formula = [f for f in cells if isinstance(f, str) and f.startswith('=')]
    for f in with_formula:
        for name in re.findall(r'\b([a-z_][a-z0-9_]*)\(', f):
            python_calls[name] = python_calls.get(name, 0) + 1
    print('  %-12s %-12s %5d formulas' % (sheet.Name, used.Address.replace('$', ''), len(with_formula)))
print()
print('Functions called, most first:')
for name, n in sorted(python_calls.items(), key=lambda kv: -kv[1])[:15]:
    print('  %-20s %d' % (name, n))
