# Python in Excel: examples

Python functions in worksheet formulas, macros that work on the workbook, scripts, and an ASF script that calls Python — all running inside Excel on your computer, with the Python the add-in carries (numpy, pandas, openpyxl and pywin32 included).

| Theme | Module | What it shows |
|---|---|---|
| Statistics | `python/xl_stats.py` | numpy on a range: z-scores, outliers, a trend line, a seasonal forecast, a moving average, a histogram |
| Tables | `python/xl_tables.py` | pandas on a table with its header row: describe, group totals, a pivot with totals, a query, a join, the top rows, a correlation matrix, and a table kept in Python as a handle |
| Finance | `python/xl_finance.py` | a loan schedule, a growth rate, and a Monte Carlo simulation that runs asynchronously while Excel goes on |
| Text | `python/xl_text.py` | the standard library alone: names tidied, accents removed, company names matched despite typos |
| Macros | `python/xl_macros.py` | Excel's object model from Python (`xl_app()`): summarize the selection, import a CSV with pandas, fit every column |
| ASF | `library/py_summary.vas` | an ASF script of the library calling Python (`py.call`), and `PY_CALL` in a cell |
| Scripts | `scripts/*.py` | for **Run Python script**: an overview of the workbook, the packages of the add-in's Python |

`PythonExamples.xlsx` has the sample data and one sheet per theme with the formulas.

## Setting up

1. **The Python add-in**, `asfpy64.xll` or `asfpy32.xll`, loaded in Excel for Windows: see [Python](../../docs/Python.md) for installing it. Its Python has every package these examples use. In 32-bit Excel the pandas examples (`xl_tables.py`, `SummarizeSelection`, `ImportCsv`) fail: pandas has no 32-bit build.
2. **Copy the files of `python/`** into the Python folder: **Python folder** on the ASF tab opens it (`%APPDATA%\ASF\python` by default). Then **Reload Python**, or restart Excel.
3. **Open `PythonExamples.xlsx`** and press Ctrl+Alt+F9 to calculate everything.
4. **For the Scripts sheet:** `ASF.xlam` loaded, `library/py_summary.vas` in the library folder (**Library folder** on the ASF tab), and the setting `PythonAccess` at `call` (**Python settings** on the ASF tab).

The functions are named `PY.something`, with `@xl_func(name='PY.TREND')`: Excel's own functions come first, so a Python function named `trend` would never be called by `=TREND(...)`. Insert Function lists them under the category **Python**, with their docstrings.

Each result that fills several cells is an array formula over its cells, which works in Excel 2016 and in Microsoft 365. In Microsoft 365 you can also write the formula in one cell and let it spill.

## The functions

**Statistics** (`xl_stats.py`)

| Formula | Gives |
|---|---|
| `=PY.ZSCORES(Data!G2:G61)` | Each value in standard deviations from the mean, one per value |
| `=PY.OUTLIERS(Data!E2:E61, 1.5)` | TRUE for each value outside Tukey's fences (k = 1.5 by default) |
| `=PY.TREND(Data!J2:J25)` | Slope, intercept and R² of the least-squares line, with headers |
| `=PY.FORECAST(Data!J2:J25, 12, 12)` | The next 12 values: the trend plus the average pattern of 12-period cycles |
| `=PY.MOVAVG(Data!J2:J25, 3)` | The moving average over 3 values |
| `=PY.HISTOGRAM(Data!E2:E61, 4)` | From, to and count of 4 equal intervals |

**Tables** (`xl_tables.py`): each takes a range with its header row, or a handle from `PY.TABLE`

| Formula | Gives |
|---|---|
| `=PY.TABLE(Data!A1:G61)` | `<DataFrame#1>`: the table kept in Python. The other functions take this cell instead of the range |
| `=PY.DESCRIBE(Data!A1:G61)` | Count, mean, spread and quartiles of each numeric column |
| `=PY.GROUP(Data!A1:G61, "Region", "Amount", "sum")` | One row per region with the sum (or `mean`, `count`, `min`, `max`, `median`) |
| `=PY.PIVOT(Data!A1:G61, "Region", "Product", "Amount")` | Amount by region and product, with totals |
| `=PY.QUERY(Data!A1:G61, "Region == ""North"" and Qty > 10")` | The rows that meet a pandas condition |
| `=PY.JOIN(Data!A1:G61, Data!L1:N5, "Region")` | Each order with the columns of the matching region |
| `=PY.TOP(Data!A1:G61, "Amount", 3)` | The 3 rows with the largest amounts |
| `=PY.CORR(Data!E1:G61)` | The correlation of every pair of numeric columns |

**Finance** (`xl_finance.py`)

| Formula | Gives |
|---|---|
| `=PY.LOAN(25000, 6.5%, 3)` | Every monthly payment: interest, principal and balance |
| `=PY.CAGR(100, 180, 5)` | The compound annual growth rate |
| `=PY.MONTECARLO(10000, 6%, 15%, 10, 20000)` | The 5th, 50th and 95th percentiles of the value after 10 years of random returns, over 20,000 runs. Asynchronous: the cell shows `#GETTING_DATA` while it runs, and Excel stays usable |

**Text** (`xl_text.py`)

| Formula | Gives |
|---|---|
| `=PY.PROPERNAME("o'brien SEAN")` | `O'Brien Sean` (also `McDonald`, `de la`, hyphenated names) |
| `=PY.NOACCENTS("José Müller")` | `Jose Muller` |
| `=PY.FUZZY("Umbrela Corp", Data!R2:R6, 0.5)` | `Umbrella Corporation`: the closest entry of the list, or empty below the cutoff |
| `=PY.SIMILARITY("Initech LLC", "initech llc")` | `1`: how alike two texts are, from 0 to 1 |

**Macros** (`xl_macros.py`): **Python macros** on the ASF tab lists and runs them; VBA runs them with `Application.Run`.

| Macro | Does |
|---|---|
| `SummarizeSelection` | Select a table with its header row: the count, mean, min and max of each numeric column are written beside it |
| `ImportCsv` | `Application.Run "ImportCsv", "C:\data\sales.csv"`: the file read with pandas into a new sheet |
| `FitAllColumns` | Fits the columns of every sheet to their contents |

**From ASF and VBA**

- `=ASF_CALL("py_summary", Data!G2:G61)` runs `library/py_summary.vas`, whose `// @asf python: call` header gets it `py`: it calls `xl_stats.summary`, marked `@export`, and builds a table from the object it returns (`s.mean`, `s.max`...).
- `=PY_CALL("xl_stats.above_mean", Data!G2:G61)` calls an exported function from a cell, without registering it as a worksheet function.

**Scripts** (`scripts/`): **Run Python script** on the ASF tab runs one and shows what it printed.

- `workbook_overview.py`: each sheet of the active workbook with its used range and number of formulas, and the functions the formulas call, most used first.
- `package_report.py`: the Python inside Excel and every package it has, with versions.

## Writing your own

- A module of the Python folder imports `from asfpy import xl_func, xl_macro, export, xl_app`.
- Type hints convert the arguments: `numpy.ndarray` (a range as an array, empty cells as `nan`), `pandas.DataFrame` (a range with its header row), `int`, `float`, `str`, `list`, `date`.
- Return a list of rows, a numpy array or a DataFrame to fill several cells; return `-> object` (or any non-cell value) to keep it in Python as a handle.
- After changing a module, **Reload Python** on the ASF tab. **Python console** on the ASF tab imports the modules of the Python folder too, to try a function before Excel loads it.

[docs/Python.md](../../docs/Python.md) describes all of it.

## Checking without Excel

```
python3 examples/python-in-excel/tools/check.py          # every function through the add-in's boot.py
python3 examples/python-in-excel/tools/make_workbook.py  # writes PythonExamples.xlsx again
```

`check.py` loads the modules through `excel/python/xll/boot.py`, as the add-in does, and calls each worksheet function as a cell calls it, with the sample data of `tools/data.py`. The macros need Excel's window, so it only checks that they are registered as commands. It needs numpy and pandas; `make_workbook.py` also openpyxl.
