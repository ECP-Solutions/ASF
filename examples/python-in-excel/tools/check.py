"""Checks the examples without Excel: the modules of python/ load through the add-in's own boot.py
(excel/python/xll/boot.py), as Excel's add-in loads them, and each worksheet function is called
as a cell calls it, with the sample data of tools/data.py; the macros and the ASF script are
checked for what can be checked outside Excel.

    python3 examples/python-in-excel/tools/check.py

Needs numpy and pandas in the Python that runs it. Prints one line per check."""
import datetime
import importlib.util
import math
import os
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
EXAMPLE = os.path.dirname(HERE)
ROOT = os.path.dirname(os.path.dirname(EXAMPLE))
sys.path.insert(0, HERE)
import data  # noqa: E402

failures = 0


def check(what, got, want):
    global failures
    ok = got == want if not callable(want) else want(got)
    failures += not ok
    print('%s %s: %s' % ('ok  ' if ok else 'FAIL', what, got if len(repr(got)) < 160 else repr(got)[:160] + '...'))


def load_boot():
    spec = importlib.util.spec_from_file_location('asfpy_boot', os.path.join(ROOT, 'excel', 'python', 'xll', 'boot.py'))
    boot = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(boot)
    saved = sys.stdout, sys.stderr
    functions = boot.start(os.path.join(EXAMPLE, 'python'), '', '', os.path.join(tempfile.gettempdir(), 'asfpy-examples.log'))
    sys.stdout, sys.stderr = saved
    return boot, [f[0] for f in functions]


def serial(d):
    return float((d - datetime.date(1899, 12, 30)).days)


def cells(boot, tagged):
    """A tagged result of boot.py as plain values: a grid as rows, a cell as its value."""
    tag, value = tagged
    if tag == boot.TAG_GRID:
        return [[cells(boot, c) for c in row] for row in value]
    if tag == boot.TAG_FAIL:
        return 'FAIL: ' + value
    return value


def main():
    boot, names = load_boot()
    info = boot.root(1).startInfo
    print(info.split('\n')[0])

    def call(name, *args):
        return cells(boot, boot.xl_call(names.index(name), list(args)))

    sales = [[serial(c) if isinstance(c, datetime.date) else c for c in row] for row in data.sales()]
    qty = [[r[4]] for r in sales[1:]]
    amount = [[r[6]] for r in sales[1:]]
    monthly = data.monthly()
    visitors = [[r[1]] for r in monthly[1:]]

    check('the modules load, nothing refused', [l for l in info.split('\n') if 'did not import' in l or 'not registered' in l], [])
    check('worksheet functions and macros registered', len(names), lambda n: n >= 20)

    # --- xl_stats: numpy
    z = call('PY.ZSCORES', amount)
    check('zscores: one per value, mean 0', (len(z), abs(sum(r[0] for r in z)) < 1e-4), (60, True))
    flags = call('PY.OUTLIERS', qty)
    check('outlier_flags: the two outsized orders', [i + 2 for i, r in enumerate(flags) if r[0]], [19, 46])
    t = call('PY.TREND', visitors)
    check('trend: header and a rising slope', (t[0], t[1][0] > 0, 0 < t[1][2] <= 1), (['slope', 'intercept', 'R2'], True, True))
    f = call('PY.FORECAST', visitors, 12, 12)
    check('forecast: 12 months, the seasonal peak in July', (len(f), max(range(12), key=lambda i: f[i][0]) + 1), (12, 7))
    ma = call('PY.MOVAVG', visitors, 3)
    check('moving_average: two empty, then the means', (ma[0][0], ma[1][0], ma[2][0] == round(sum(v[0] for v in visitors[:3]) / 3, 6)), (None, None, True))
    h = call('PY.HISTOGRAM', qty, 4)
    check('histogram: 4 intervals counting every order', (len(h), sum(r[2] for r in h[1:])), (5, 60))

    # --- xl_tables: pandas
    s = call('PY.DESCRIBE', sales)
    check('table_summary: describe of the numeric columns', (s[0][1:], s[1][0]), (['Date', 'Qty', 'Price', 'Amount'], 'count'))
    g = call('PY.GROUP', sales, 'Region', 'Amount')
    check('group_total: one row per region, largest first', (g[0], len(g), g[1][1] >= g[2][1]), (['Region', 'Amount'], 5, True))
    total = round(sum(r[6] for r in sales[1:]), 2)
    check('group_total: the regions add up to the total', round(sum(r[1] for r in g[1:]), 2), total)
    p = call('PY.PIVOT', sales, 'Region', 'Product', 'Amount')
    check('pivot: regions by products, with totals', (p[0], p[-1][0], round(p[-1][-1], 2)),
          (['Region', 'Gadget', 'Gizmo', 'Widget', 'Total'], 'Total', total))
    j = call('PY.JOIN', sales, data.regions(), 'Region')
    check('join_tables: each order with its manager', (j[0][-2:], len(j)), (['Manager', 'Target'], 61))
    q = call('PY.QUERY', sales, 'Region == "North" and Qty > 10')
    check('query: northern orders over 10', all(r[2] == 'North' and r[4] > 10 for r in q[1:]) and len(q) > 1, True)
    top = call('PY.TOP', sales, 'Amount', 3)
    check('top_n: the 3 largest orders, the two outsized first', sorted(r[0] for r in top[1:3]), ['SO-1018', 'SO-1045'])
    c = call('PY.CORR', [[r[4], r[5], r[6]] for r in sales])
    check('correlation: 3 by 3, 1 on the diagonal', (len(c), c[1][1], c[2][2]), (4, 1.0, 1.0))

    boot._native_caller = lambda: '[PythonExamples.xlsx]Tables!R2C10'
    handle = call('PY.TABLE', sales)
    check('load_table: a handle', handle, lambda h: isinstance(h, str) and h.startswith('<DataFrame#'))
    boot._native_caller = lambda: '[PythonExamples.xlsx]Tables!R2C12'
    check('a handle in another function', call('PY.GROUP', handle, 'Product', 'Qty', 'count')[0], ['Product', 'Qty'])
    boot._native_caller = None

    # --- xl_finance
    sched = call('PY.LOAN', 10000, 0.06, 1)
    check('loan_schedule: 12 payments, paid off', (len(sched), sched[1][1], sched[-1][4]), (13, 860.66, 0.0))
    check('cagr: 100 to 200 in 5 years', round(call('PY.CAGR', 100, 200, 5), 6), 0.148698)
    mc = call('PY.MONTECARLO', 10000, 0.06, 0.15, 10, 5000, 1)
    check('monte_carlo: P5 < P50 < P95', (mc[0], mc[1][0] < mc[1][1] < mc[1][2]), (['P5', 'P50', 'P95'], True))

    # --- xl_text
    typed, clean = data.customers()
    check('fuzzy_match: a typo', call('PY.FUZZY', 'Umbrela Corp', [[c] for c in clean]), 'Umbrella Corporation')
    check('fuzzy_match: nothing close', call('PY.FUZZY', 'zzz', [[c] for c in clean]), '')
    check('similarity', call('PY.SIMILARITY', 'Initech LLC', 'initech llc'), 1.0)
    check('strip_accents', call('PY.NOACCENTS', 'José Müller'), 'Jose Muller')
    check('proper_name', [call('PY.PROPERNAME', t) for t in ('  josé  müller', "o'brien SEAN", 'Ana de la ROSA', 'mcdonald ronald')],
          ['José Müller', "O'Brien Sean", 'Ana de la Rosa', 'McDonald Ronald'])

    # --- for ASF scripts and PY_CALL
    exported = boot._exports
    s = exported['xl_stats.summary']([r[0] for r in amount])
    check('xl_stats.summary for py.call: an object', (s.count, round(s.mean, 2) == round(total / 60, 2), s.max > s.min), (60, True, True))
    check('PY_CALL("xl_stats.above_mean", ...)', cells(boot, boot.xl_call_export('xl_stats.above_mean', [amount])),
          lambda n: 0 < n < 60)
    vas = open(os.path.join(EXAMPLE, 'library', 'py_summary.vas'), encoding='utf-8').read()
    check('py_summary.vas asks for python: call', '// @asf python: call' in vas, True)

    # --- scripts for Run in Excel: compiled; the package report run as the add-in runs it
    for name in ('workbook_overview.py', 'package_report.py'):
        path = os.path.join(EXAMPLE, 'scripts', name)
        compile(open(path, encoding='utf-8').read(), path, 'exec')
    out = cells(boot, boot.run_script(os.path.join(EXAMPLE, 'scripts', 'package_report.py')))
    check('Run in Excel: package_report.py', out.split('\n')[0], lambda line: line.startswith('Python 3.'))
    check('  it lists numpy and pandas', ('numpy ' in out, 'pandas ' in out), (True, True))

    # --- macros: registered as commands; they need Excel's window (xl_app), so not run here
    by_name = {f[0]: f for f in boot._functions}
    check('macros registered as commands', sorted(f[2] for f in by_name.values() if f[3] == '%'),
          sorted(['SummarizeSelection', 'ImportCsv', 'FitAllColumns']))

    print('%d checks failed' % failures if failures else 'all checks pass')
    return 1 if failures else 0


if __name__ == '__main__':
    sys.exit(main())
