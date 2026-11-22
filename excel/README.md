# ASF add-in for Excel

`ASF.xlam` brings ASF to Excel as an add-in: one file, installed once, built from `src/`. Plain `.xlsx` workbooks can then run ASF scripts in worksheet functions, without macros of their own. It is the Excel counterpart of the LibreOffice extension (`libreoffice/ASF.oxt`).

It works in Excel 2016 and later, on Windows and Mac. Excel for the web, iPad and Teams don't run VBA, so workbooks that use the functions show `#NAME?` there.

## Worksheet functions

| Function | What it does |
|---|---|
| `ASF_RUN(script, [arg1], …)` | Runs `script`, with the arguments as `$1`, `$2`… |
| `ASF_CALL(name, [arg1], …)` | Runs `name.vas` from the script library folder |
| `RX_TEST(text, pattern, [ignoreCase])` | `TRUE` when the text matches the regular expression |
| `RX_EXTRACT(text, pattern, [group], [ignoreCase])` | The first match, or one of its groups |
| `RX_REPLACE(text, pattern, replacement, [ignoreCase])` | Every match replaced (`$1`, `$2`… for groups) |
| `RX_COUNT(text, pattern, [ignoreCase])` | The number of matches |

```
=ASF_RUN("return $1 * 1.16;", B2)
=ASF_RUN("return $1.filter(fun(r) { return r[3] > 100; });", A2:D500)
=ASF_RUN(H1, A2:D500)                      the script in cell H1
=ASF_CALL("discounts", Orders!A1:F200, Rules!A1:C20)
=RX_EXTRACT(A2, "(\d+)\s*kg", 1)
```

A string in an Excel formula can be at most 255 characters long. Put a longer script in a cell and pass the cell, or put it in the library and use `ASF_CALL`.

The regular expressions are ASF's own engine (JavaScript syntax), so results are the same on Windows and Mac.

### Writing formulas with them

The functions are registered with Excel, with a description for each one and for each of its arguments, in the category **ASF**:

- **Insert Function (fx)** lists them under ASF, with their descriptions.
- **Ctrl+A** after `=ASF_RUN` opens the Function Arguments dialog: a box per argument, each with its description, and the result as you fill them.
- **Ctrl+Shift+A** after the name types the argument names in: `=ASF_RUN(script,args)`.
- **Formula AutoComplete** offers the names as you type `=ASF_` or `=RX_`.

The tips Excel shows under the formula while you type, as for `SUM(number1, [number2], ...)`, are only for Excel's own functions. The free [Excel-DNA IntelliSense](https://github.com/Excel-DNA/IntelliSense) add-in shows them for these functions too. Load its `.xll` next to `ASF.xlam` and it reads the tips from the add-in's `_IntelliSense_` sheet:

```
ASF_RUN(script, [args]...)
RX_EXTRACT(text, pattern, [group], [ignoreCase])
```

### Arguments

- A single cell gives its value. Numbers become doubles, as numbers in scripts are.
- One column or one row of cells gives a flat array from 1: with `B2:B6`, `$1[1]` is B2 and `$1.filter(fun(x) { return x > 2; })` works on the values.
- A range of several rows and columns gives an array of rows, each row an array from 1: `$1[row][column]`.
- An empty cell is `undefined`; scripts write `x ?? ''` when they want text.
- An error value as a whole argument (`#N/A` in a single cell) is returned as it is, without running the script, as Excel's own functions do.
- Errors inside a range arrive as their text (`"#N/A"`), so a script can test for them.
- Ranges are passed as values, never as `Range` objects.

### Results

| The script returns | The cell shows |
|---|---|
| a number, string or boolean | the value |
| `undefined` or `null` | an empty string |
| an array of rows | a grid |
| a flat array | one column |
| an object, or an array inside a cell | its text, as `print` writes it |
| an error, an uncaught `throw`, a compile error, or the step limit | `#VALUE!` |

How array results reach the cells depends on the version:

- **Microsoft 365 and Excel 2021:** the result spills from the formula's cell.
- **Excel 2016 and 2019:** select the output range, type the formula and press Ctrl+Shift+Enter. The result is fitted to the selection: cells beyond it get `""`, and a larger result is cut.
- The same workbook works in both.

## Example: a data pipeline

[examples/excel-pipeline](../examples/excel-pipeline/README.md) is a workbook that only uses formulas. Raw orders are cleaned, checked against rules kept in a sheet, de-duplicated, enriched from a reference table and summarized, one `ASF_CALL` per stage, with its scripts in the library folder.

## Script library and settings

`ASF_CALL` runs scripts from the library folder. That folder is also the only place `import` may read from, for `ASF_CALL` and `ASF_RUN` alike.

| Setting | Default |
|---|---|
| `LibraryFolder` | Windows: `%APPDATA%\ASF\scripts`. Mac: `~/Library/Group Containers/UBF8T346G9.Office/ASF/scripts`, where Excel for Mac can read without asking |
| `RunEnabled` | `1`. With `0`, `ASF_RUN` gives `#VALUE!` and only library scripts run |
| `StepLimit` | `10000000` loop iterations and calls per script; `0` for no limit |
| `PythonAccess` | `off`. What library scripts that ask for Python may do: `call` (`py.call` of exported functions) or `modules` (`py.module` too). See **Python** below |
| `PythonVenv`, `PythonFolder` | Read by the Python add-in when Excel loads it: another Python than the one it carries (a venv or an install; empty by default), and its folder of modules (default `%APPDATA%\ASF\python`) |

The settings are kept with `SaveSetting`: on Windows in `HKEY_CURRENT_USER\Software\VB and VBA Program Settings\ASF\AddIn`, where an administrator can set them by group policy; on Mac in a preferences file. Change one from the Immediate window or a macro:

```vb
Application.Run "ASF.xlam!ASFSettings.ChangeSetting", "RunEnabled", "0"
```

The worksheet functions keep one engine and its compiled scripts for the whole session. A library script is read again when its file changes. Scripts are read as ANSI text, so save them as ASCII or Windows-1252.

## The ASF tab

| Button | What it does |
|---|---|
| **Run script** | Lists the scripts of the library folder (the first 30) and runs the one you pick, then shows its result or why it stopped |
| **Library folder** | Opens the library folder, made first when it is missing |
| **Why this error?** | Calculates the active cell again and shows why its `ASF_RUN` or `ASF_CALL` gave `#VALUE!` (with several calls in one formula, the message of the last one). With the Python add-in loaded, also the Python error and traceback of a `PY_CALL` or Python function |
| **Reset engine** | Drops the engine of the worksheet functions and their compiled scripts; the next formula reads the settings again |
| **Run Python script** | Runs a `.py` file you pick in Excel's Python (the Python add-in's command `AsfPyRunScript`) and shows what it printed |
| **Python macros** | Lists the `@xl_macro` functions of the Python folder (`AsfPyMacros`) and runs the one you pick, asking for its arguments |
| **Reload Python** | The Python add-in imports its folder again and registers its functions again, then the workbooks are calculated in full |
| **Python settings** | Which Python the add-in runs (its own, or a venv or install), the Python folder, and `PythonAccess`, each shown and changed in turn |
| **Python packages** | A console window with pip for the Python add-in's Python (`install numpy`); Excel uses new packages after a restart |
| **Python console** | The Python add-in's Python at its `>>>` prompt, in a console window apart from Excel (`AsfPyConsole`): `pip("install pandas")`, `import pandas`, and the Python folder's modules to try |
| **Python folder** | Opens the Python add-in's folder of modules, made first when it is missing |
| **About Python** | The loaded add-in file, its Python, the settings, the Python folder's start-up and the log |
| **Help** | Opens this page |

A library script may start with a header of comment lines. It says how **Run script** lists and runs it; the worksheet functions ignore it.

```
// @asf name: Apply discount rules
// @asf appAccess: true
// @asf args: selection
// @asf python: call
```

| Key | Meaning |
|---|---|
| `name` | The name the list shows; the file name when there is none |
| `appAccess` | `true`: the script may use Office objects. Default `false` |
| `args` | `selection`: `$1` is the selected range. Default: no arguments |
| `python` | `call` or `modules`: the script gets `py`, from the Python add-in. Unlike the other keys, `ASF_CALL` reads it too. See **Python** below |

With `appAccess: true` and `args: selection`, `$1` is the `Range` itself, so the script can write cells, format them or reach the workbook. Without `appAccess`, `$1` holds the range's values, as in a formula. Only library scripts can ask for Office access: the library holds what the user installed.

## Security

A workbook's formulas run when it recalculates, with no macro warning: the add-in is trusted, and the scripts are only data. So every script the functions run is limited, wherever it comes from:

- **No Office objects.** `AppAccess` is off and ranges arrive as values.
- **Imports only from the library folder** (`ImportRoot`). An import with `../` or an absolute path outside it gives `#VALUE!`.
- **A step limit** (`MaxSteps`). `while (true) { }` in a cell can't freeze Excel.
- **`ASF_CALL` only runs installed scripts.** It only takes bare script names, so a received workbook can only run scripts the user installed in the library.
- **Scripts are never stored inside workbooks.**
- **`ASF_RUN` can be turned off** in the settings, leaving only library scripts.
- **Office access only for library scripts run from the ribbon,** and only when their header asks for it.
- **Python only for library scripts, with two keys:** the script's `python` header and the `PythonAccess` setting. A formula's script (`ASF_RUN`) never gets `py`.

## Python

With the Python add-in (`asfpy64.xll` or `asfpy32.xll`, Excel for Windows; see `excel/python`) loaded next to `ASF.xlam`, a library script can call Python. It asks with its header, and the `PythonAccess` setting decides:

```js
// @asf python: call
return py.call('stats.zscore', $1);          // stats.py in the Python folder, a function marked @export
```

```js
// @asf python: modules
let np = py.module('numpy');
return np.percentile($1, 90);
```

| Header | `PythonAccess` | The script |
|---|---|---|
| none | any | runs without `py` |
| `call` | `off` (the default) | doesn't run: `#VALUE!`, and **Why this error?** names the setting |
| `call` | `call` or `modules` | gets `py`: `py.call` of exported functions |
| `modules` | `off` or `call` | doesn't run, as above |
| `modules` | `modules` | gets `py`: `py.call` and `py.module` |

```vb
Application.Run "ASF.xlam!ASFSettings.ChangeSetting", "PythonAccess", "call"
```

A script that asks for more than the setting allows doesn't run at all, so it never fails halfway through. A Python error stops the script; **Why this error?** shows the script's message and Python's. Python can't be sandboxed: `modules` lets library scripts do anything Python can, which is why it is a setting of its own. `PY_CALL` and the Python folder's worksheet functions come from the Python add-in itself and don't depend on `PythonAccess`.

### Installing the Python add-in

The add-in carries its own Python (CPython 3.14 of its bitness, with pip): nothing else to install, no venv to make.

1. Unblock `asfpy64.xll` (64-bit Excel) or `asfpy32.xll` (32-bit Excel; File > Account > About Excel says which) in its Properties, copy it to a trusted location, and add it under File > Options > Add-ins > Manage: Excel Add-ins > Browse. The first time, it unpacks its Python into `%LOCALAPPDATA%\ASF\runtime` (about 40 MB).
2. Put your modules in the Python folder, `%APPDATA%\ASF\python` (**Python folder** on the ASF tab opens it).
3. For packages, **Python packages** on the ASF tab (or `Application.Run "AsfPyPackages"`) opens a console window with pip for that Python: type `install numpy pandas pywin32`, then restart Excel. Close Excel before upgrading or removing a package it has loaded.

To use another Python instead, such as a venv with pinned packages: a python.org Python 3.11 or later of Excel's bitness (not the Microsoft Store Python, nor a free-threaded build), set with the Immediate window (Ctrl+G) and `ASF.xlam` loaded:
```vb
Application.Run "ASF.xlam!ASFSettings.ChangeSetting", "PythonVenv", Environ$("USERPROFILE") & "\asfpy-venv"
```
Packages then go into that venv, with its own pip; **Python packages** runs its `python.exe`.

Python starts once per Excel: after changing the settings or the add-in, restart Excel. After changing a module, **Reload Python** is enough. If Python doesn't start, Excel says why when it loads the add-in, and `%TEMP%\asfpy.log` has the details. `excel/python/tests` holds the add-in's tests, and the checklist to run in Excel.

### The Python folder's functions

The modules of the Python folder mark what Excel gets:

```python
from asfpy import export, xl_app, xl_func, xl_macro
import numpy, pandas

@xl_func                                  # =zscore(A2:A500), a worksheet function
def zscore(values: numpy.ndarray):
    return (values - values.mean()) / values.std()

@xl_func
def summary(table: pandas.DataFrame):     # the range's first row holds the column names
    return table.describe()               # back with its column names and index

@xl_func
async def quote(symbol):                  # asynchronous: Excel goes on with other cells meanwhile
    ...

@xl_macro                                 # Application.Run "tidy", "Sheet1"
def tidy(sheet):
    xl_app().Worksheets(sheet).UsedRange.Columns.AutoFit()   # needs pywin32 (Python packages)

@export                                   # =PY_CALL("stats.scale", A1:A3, 10), py.call in scripts
def scale(values: list, factor: float):
    return [v * factor for v in values]
```

- **Type hints** convert the arguments: `int`, `float`, `str`, `bool`, `list`, `date`, `datetime`, `numpy.ndarray` (numbers as `float`, empty cells as `nan`; one row or column 1-D), `pandas.DataFrame` (the first row gives the column names) and `pandas.Series`. Without a hint a range is a list of rows, one row or column a flat list.
- **Results:** a list of rows, a numpy array or a DataFrame fills the cells of an array formula (Excel 365 spills it). A DataFrame comes with its column names, and with its index when the index has a name or isn't whole numbers. Missing values (`NaN`, `NaT`, `None`) are empty cells; `nan` alone is `#NUM!`.
- **Objects:** any other result, or any result of a function hinted `-> object`, stays in Python, and the cell shows a handle such as `<Counter#12>`. Another function given that cell gets the object. The handle belongs to its cell: calculated again, the cell gets a new object and the old one goes. A handle names nothing in another workbook, or after Excel restarts (calculate the cell again).
- **Asynchronous functions** (`async def`, or `@xl_func(asynchronous=True)` for a plain function on a thread of the add-in) show `#GETTING_DATA` until they are done; Excel's calculation waits for them.
- **Commands** (`@xl_macro`) run with `Application.Run "name", arguments...` and give back their result. They run on Excel's main thread, and only they may use `xl_app()`, Excel's `Application` (with the package pywin32 or comtypes).
- `thread_safe=True` lets Excel call a function on its calculation threads, `volatile=True` calculates it on every recalculation.

## Macros in other workbooks

Add a reference to the add-in (VBA editor: Tools > References > ASFAddIn), then:

```vb
Dim e As ASF
Set e = ASFAddIn.NewEngine()
Debug.Print e.Run(e.Compile("return 6 * 7;"))
```

`NewEngine` returns a new engine with `AppAccess` off, separate from the one the worksheet functions share. Importing the modules of `src/` into a workbook still works as before; the add-in is an alternative.

## Install

The add-in is not signed. Install it in a trusted location.

**Windows**
1. If the file was downloaded, unblock it: right-click it, then Properties > Unblock.
2. Copy it to `%APPDATA%\Microsoft\AddIns`, which Excel trusts by default, or to a folder added under Trust Center > Trusted Locations.
3. In Excel: File > Options > Add-ins > Manage: Excel Add-ins > Go, then tick ASF (or Browse to the file).

**Mac**
1. In Excel: Tools > Excel Add-ins > Browse, and select `ASF.xlam`.
2. Enable its macros when Excel asks.

An organisation that only allows signed VBA can sign the project with its own certificate (VBA editor: Tools > Digital Signature) before deploying it.

## Build

On Windows with Excel, from the repository:

```powershell
powershell -ExecutionPolicy Bypass -File excel/tools/build.ps1 -TemporaryTrustAccess
```

The script:
1. Imports the twelve classes of `src/`. They are made visible to referencing projects (`VB_Exposed = True`), so a workbook can declare `As ASF`.
2. Imports the modules of `excel/addin` and `examples/use-cases/regex/RegexFunctions.bas`.
3. Names the project `ASFAddIn` and registers the function descriptions (category ASF in Insert Function).
4. Saves `excel/ASF.xlam`.
5. Adds the ASF tab, `excel/customUI/customUI14.xml`, to the saved file.

The descriptions of the functions and their arguments come from one table, `ASFRegister.FunctionInfo`, used in two places:
- **Registration:** the build registers them, and the code `excel/addin/ThisWorkbook.vba` adds to the add-in registers them again on every opening, because Excel doesn't keep the argument descriptions in the file.
- **Typing tips:** the build also writes the table to the `_IntelliSense_` sheet for Excel-DNA IntelliSense.

Excel must trust access to the VBA project object model. `-TemporaryTrustAccess` turns it on for the build and back as it was.

| Module | Holds |
|---|---|
| `ASFFunctions` | `ASF_RUN`, `ASF_CALL`, `NewEngine` |
| `ASFEngine` | the engine the functions share |
| `ASFCells` | the conversions between cells and ASF values |
| `ASFSettings` | the settings |
| `ASFRegister` | the function descriptions: Insert Function, the Function Arguments dialog, the `_IntelliSense_` sheet |
| `ASFLibrary` | the scripts of the library and their headers, run as the ribbon runs them |
| `ASFRibbon` | the callbacks of the ASF tab |

`ASFFunctions` and `ASFRibbon` are the only modules other code reaches. The rest are `Option Private Module`, so their functions don't show up as worksheet functions. `ASFRibbon` holds only Subs, which never do.

## Tests

- **`python3 excel/tools/run_lo.py`** runs every case of `excel/tools/cases.py` in LibreOffice, headless. It calls the add-in's modules from Basic, with a temporary library folder and a low step limit. It also checks:
  - the library as **Run script** lists and runs it;
  - the script headers;
  - that every button of the ribbon calls a Sub of `ASFRibbon`;
  - that `ASFRegister` describes every worksheet function, with its arguments in the order of its parameters.

  It needs LibreOffice and `python3-uno`.
- **The buttons themselves** need Excel: try each once after a build.
- **`excel/test/AddinTests.xlsx`** holds the same cases as formulas, except those that need the library folder or a low step limit. Open it with the add-in loaded: B1 says whether every check passes. Array results are array formulas, as Excel 2016 enters them.
- **Spilling** in Microsoft 365 is checked by hand: `=ASF_RUN("return [[1, 2], [3, 4]];")` in one cell should fill two by two.
- **`python3 excel/tools/make_workbook.py`** writes the workbook again after the cases change. It needs `openpyxl`.
