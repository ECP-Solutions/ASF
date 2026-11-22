---
layout: default
title: Python
has_children: false
nav_order: 6
description: "The ASF Python add-in for Excel: installation, Python functions in cells, PY_CALL, packages, and Python in ASF scripts."
---

# Python in Excel
{: .no_toc }

The Python add-in runs Python inside Excel for Windows, on your computer:

- functions you write in Python become **worksheet functions**, with numpy and pandas, objects kept between cells, asynchronous functions and commands;
- `PY_CALL("module.function", ...)` calls a Python function from a cell;
- ASF scripts of your library can call Python through `py`;
- VBA can call it too.

It is a separate add-in from [`ASF.xlam`](Add-ins.html), which keeps working without Python. It comes as two files: `asfpy64.xll` for 64-bit Excel and `asfpy32.xll` for 32-bit Excel. Each carries its own Python (CPython 3.14, with pip), so there is nothing else to install. Unlike Excel's own `=PY()`, which runs in Microsoft's cloud, everything runs on your computer and works offline.

1. TOC
{:toc}

---

## Installation

Testing a beta? Follow [Beta testing](Beta%20testing.html) instead: it starts with cleaning up earlier test runs.

You need Excel 2016 or later, or Microsoft 365, for Windows. Excel for Mac, the web and iPad can't load `.xll` add-ins.

### 1. Find Excel's bitness

In Excel: File > Account > **About Excel**. The first line ends with **64-bit** or **32-bit**. Most Microsoft 365 installs are 64-bit; many Excel 2016 installs are 32-bit.

### 2. Download the add-in

From the [latest release](https://github.com/ECP-Solutions/ASF/releases/latest), download the file for your Excel:

| Excel | File |
|---|---|
| 64-bit | [`asfpy64.xll`](https://github.com/ECP-Solutions/ASF/releases/latest/download/asfpy64.xll) |
| 32-bit | [`asfpy32.xll`](https://github.com/ECP-Solutions/ASF/releases/latest/download/asfpy32.xll) |

The other one doesn't load in your Excel.

{: .note }
The add-in files of v4.1.0 don't carry Python yet: with them, install a Python and set it as described in [Using another Python](#using-another-python).

### 3. Unblock it and put it in a trusted location

Excel refuses add-ins downloaded from the internet.

1. Right-click the file, choose **Properties**, tick **Unblock** and click **OK**. Windows doesn't show the box when there is nothing to unblock.
2. Copy it to a folder of your own that stays: `%APPDATA%\Microsoft\AddIns` (type the path in File Explorer's address bar) is trusted by Excel by default.

### 4. Load it

In Excel: File > Options > Add-ins. At the bottom, choose **Manage: Excel Add-ins** and click **Go**, then **Browse**, pick the `.xll` and click **OK**. Make sure it is ticked.

The first time, the add-in unpacks its Python into `%LOCALAPPDATA%\ASF\runtime` (about 40 MB, once). Excel then loads it every time it starts.

If Python can't start, Excel says why right away. The details are in `%TEMP%\asfpy.log`. See [Troubleshooting](#troubleshooting).

### 5. Your first function

The add-in reads the Python files of the **Python folder**, `%APPDATA%\ASF\python`. Create the folder, and in it a file `hello.py`:

```python
from asfpy import xl_func

@xl_func
def hello(name):
    """Says hello."""
    return 'Hello, ' + name
```

Restart Excel (or, with `ASF.xlam`, click **Reload Python** on the ASF tab). In a cell, `=hello("world")` gives `Hello, world`. Until a file of the folder gives a function, Excel shows a message when it loads the add-in.

### 6. Packages

The add-in's Python has the standard library and pip. To add numpy, pandas or any other package:

1. Open the package console: **Python packages** on the ASF tab of `ASF.xlam`, or in the Immediate window of the VBA editor (Alt+F11, then Ctrl+G):
   ```vb
   Application.Run "AsfPyPackages"
   ```
2. A console window opens with the packages installed so far. Type a pip command and press Enter:
   ```
   pip> install numpy pandas pywin32
   ```
   Other commands: `install --upgrade pandas`, `uninstall -y openpyxl`, `list`.
3. Wait for `pip exited with 0`, then press Enter on an empty line to close the window.
4. **Restart Excel** to use the new packages.

pip downloads from PyPI. Behind a company proxy or offline, pip takes the usual options: `install C:\downloads\some_package.whl`, or `install --index-url https://your.mirror/simple numpy`.

Windows locks the files of a package that Excel has loaded: close Excel before you upgrade or remove such a package, and run the console from a new Excel without using Python (or with the add-in unticked).

### The Python console

**Python console** on the ASF tab (or `Application.Run "AsfPyConsole"`) opens a console window with the add-in's Python at its `>>>` prompt. It runs apart from Excel, so nothing typed there can break Excel's Python. Use it to manage packages and try them:

```python
>>> pip("install pandas")          # pip for this Python: install, list, show, uninstall -y ...
>>> pip("list")
>>> import pandas; pandas.__version__   # the version Excel will load after a restart
>>> import stats                   # a module of your Python folder, before Excel loads it
>>> stats.zscore([1, 2, 3])
```

The Python folder is on the import path, and `from asfpy import xl_func` works there: the decorators mark nothing, and `xl_app()` says it works inside Excel only. As with the package console, Excel uses new packages after a restart. `exit()` closes the window.

### 7. ASF.xlam (optional)

With [`ASF.xlam`](Add-ins.html) loaded too, the ASF tab gets two Python groups, ASF scripts of your library can call Python, and **Why this error?** shows Python's error and traceback.

**Python**

| Button | What it does |
|---|---|
| **Run Python script** | Pick a `.py` file: it runs in Excel's Python as `__main__`, with its folder on the import path, and what it prints is shown (in Notepad when it's long). Like a macro, it may use Excel through `xl_app()` |
| **Python macros** | Lists the `@xl_macro` functions of the Python folder and runs the one you pick, asking for its arguments (up to 5, typed separated by commas) |
| **Reload Python** | Imports the Python folder again and registers its functions, then calculates the workbooks in full |

**Python setup**

| Button | What it does |
|---|---|
| **Python settings** | Shows and changes the three settings: which Python (the add-in's own, or a venv or install you pick), the Python folder, and `PythonAccess`. Restart Excel after changing the first two |
| **Python packages** | Opens the package console |
| **Python console** | Opens the add-in's Python at its `>>>` prompt, in a console window apart from Excel (below) |
| **Python folder** | Opens the Python folder, made first when it is missing |
| **About Python** | The loaded add-in file, the Python it runs, the settings, what the Python folder gave at start-up, and the log, which it offers to open. The first place to look when something doesn't work |

### Updating and removing

- **Update:** close Excel, replace the `.xll`, start Excel. A new add-in with a newer Python unpacks it into a new folder, which starts without packages: install them again. (Before updating, `list` in the package console shows what you had.)
- **Remove:** untick the add-in in File > Options > Add-ins (Excel Add-ins), close Excel and delete the `.xll`. Delete `%LOCALAPPDATA%\ASF\runtime` to remove the Python and its packages. Your Python folder stays.

---

## Writing functions

Every `.py` file of the Python folder is imported when Excel loads the add-in. Decorators from the `asfpy` module mark what Excel gets:

```python
from asfpy import export, xl_app, xl_func, xl_macro
import numpy, pandas

@xl_func                                  # =zscore(A2:A500)
def zscore(values: numpy.ndarray):
    """Standard scores of the values."""
    return (values - values.mean()) / values.std()

@xl_func
def summary(table: pandas.DataFrame):     # the range's first row holds the column names
    return table.describe()               # back with its column names and index

@xl_func
async def quote(symbol):                  # asynchronous: Excel goes on with other cells meanwhile
    ...

@xl_macro                                 # Application.Run "tidy", "Sheet1"
def tidy(sheet):
    xl_app().Worksheets(sheet).UsedRange.Columns.AutoFit()   # needs the package pywin32

@export                                   # =PY_CALL("stats.scale", A1:A3, 10)
def scale(values: list, factor: float):
    return [v * factor for v in values]
```

| Decorator | Makes |
|---|---|
| `@xl_func` | A worksheet function named as the Python function, in the **Python** category of Insert Function, with its docstring as the description. Options: `name=`, `volatile=True` (calculated on every recalculation), `thread_safe=True` (may run on Excel's calculation threads), `asynchronous=True` (a plain function run on a thread of the add-in; `async def` is asynchronous by itself) |
| `@xl_macro` | A command: `Application.Run "name", arguments...` runs it and gives back its result. It runs on Excel's main thread, and only commands may use `xl_app()`, Excel's `Application` (needs the package `pywin32` or `comtypes`). Option: `name=` |
| `@export` | A function `PY_CALL` and ASF scripts may call, as `"module.function"`, without making it a worksheet function. They may call `@xl_func` and `@xl_macro` functions the same way |

A function takes at most 32 arguments; the add-in registers up to 2,000 functions.

### Arguments

- **Without a type hint:** a cell gives its value (`float`, `str`, `bool`, or `None` when empty); a range gives a list of rows, and one row or one column a flat list.
- **Type hints** convert them: `int`, `float`, `str`, `bool`, `list`, `date`, `datetime`, `numpy.ndarray` (numbers as `float`, empty cells as `nan`; one row or column 1-D), `pandas.DataFrame` (the first row gives the column names) and `pandas.Series`.
- An error value as a whole argument is the result, without calling Python, as Excel's own functions do.

### Results

| The function returns | The cell shows |
|---|---|
| a number, string, boolean, date | the value |
| `None` | an empty string |
| a list of rows, a numpy array, a DataFrame, a Series | the cells of an array formula (Microsoft 365 spills it); a DataFrame with its column names, and with its index when it has a name or isn't whole numbers |
| `nan` or infinity alone | `#NUM!` |
| any other object, or anything from a function hinted `-> object` | a handle such as `<Counter#12>` (below) |
| an exception | `#VALUE!`; **Why this error?** on the ASF tab shows the traceback |

**Handles.** An object that isn't a cell value stays in Python, and its cell shows a handle. Another function given that cell gets the object itself. The handle belongs to its cell: calculated again, the cell gets a new object and the old one goes. A handle names nothing in another workbook, or after Excel restarts: calculate the cell again.

**Asynchronous functions** show `#GETTING_DATA` until they are done; Excel's calculation waits for them.

### Changing your code

Python starts once per Excel session. After changing a file of the Python folder, **Reload Python** on the ASF tab (or `Application.Run "AsfPyReload"`) imports the folder again and registers its functions; then press Ctrl+Alt+F9 to calculate everything again. Packages stay loaded: after installing or upgrading one, restart Excel.

---

## PY_CALL

`PY_CALL("module.function", arguments...)` calls a function of the Python folder by its module and name: one marked with `@export`, or with `@xl_func` or `@xl_macro`:

```
=PY_CALL("stats.scale", A1:A3, 10)
```

It takes up to 29 arguments. It only finds the functions of the Python folder marked that way: any other name (such as `"os.system"`) gives `#NAME?`, and it never imports anything. A Python exception gives `#VALUE!`.

`PY_CALL` and the `@xl_func` functions work whenever the Python add-in is loaded. Loading it is your decision, and a workbook can only pick which of your installed functions runs, with what data. So write exported functions to receive untrusted data: no paths, URLs, `eval` or pickle taken from their arguments.

---

## Python in ASF scripts

With `ASF.xlam` loaded too, a script of your [library](Add-ins.html#the-script-library) can call Python through `py`. It asks with its header, and the setting `PythonAccess` decides:

```js
// @asf python: call
return py.call('stats.zscore', $1);          // a function of stats.py marked @export
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
| `call` | `call` or `modules` | gets `py.call` of exported functions |
| `modules` | `off` or `call` | doesn't run |
| `modules` | `modules` | gets `py.call` and `py.module`, which imports any module |

Turn it on from the Immediate window (Alt+F11, then Ctrl+G):

```vb
Application.Run "ASF.xlam!ASFSettings.ChangeSetting", "PythonAccess", "call"
```

Python can't be sandboxed: `modules` lets library scripts do anything Python can, which is why it is a setting of its own. A formula's script (`ASF_RUN`) never gets `py`. A script that asks for more than the setting allows doesn't run at all, so it never fails halfway through.

---

## From VBA

Commands run with `Application.Run`, the Python folder's and the add-in's own:

```vb
Debug.Print Application.Run("tidy", "Sheet1")
Debug.Print Application.Run("AsfPyRunScript", "C:\work\report.py")   ' what the script printed
Debug.Print Application.Run("AsfPyMacros")       ' a line per @xl_macro: name, arguments it needs, description
Application.Run "AsfPyReload"                    ' the Python folder again
Application.Run "AsfPyPackages"                  ' the package console
```

For `py`, declare the add-in's function and ask for a root object: level 1 has `py.call`, level 2 also `py.module`.

```vb
#If Win64 Then
Private Declare PtrSafe Function AsfPy_Root Lib "asfpy64.xll" (ByVal level As Long, ByRef out As Variant) As Long
#Else
Private Declare PtrSafe Function AsfPy_Root Lib "asfpy32.xll" (ByVal level As Long, ByRef out As Variant) As Long
#End If

Sub Demo()
    Dim py As Variant, v As Variant
    AsfPy_Root 1, py                                         ' error 53 when the add-in isn't loaded
    Debug.Print CallByName(py, "call", VbMethod, "hello.hello", "VBA")   ' Hello, VBA
    v = CallByName(py, "call", VbMethod, "stats.scale", Array(1, 2, 3), 10)
    Debug.Print v(0)                                         ' a list comes back as an array from 0
End Sub
```

A Python exception raises a VBA error; `CallByName(py, "lastError", VbGet)` has Python's message.

---

## Settings

The add-in reads its settings when Excel loads it, from `HKEY_CURRENT_USER\Software\VB and VBA Program Settings\ASF\AddIn`. After changing one, restart Excel.

| Setting | Default |
|---|---|
| `PythonFolder` | `%APPDATA%\ASF\python`: the folder of your Python files |
| `PythonVenv` | empty: the Python the add-in carries. Otherwise, another Python to use (below) |
| `PythonAccess` | `off`: read by `ASF.xlam`, for ASF scripts (above) |

With `ASF.xlam` loaded, **Python settings** on the ASF tab shows and changes them. Or set them from the Immediate window:

```vb
Application.Run "ASF.xlam!ASFSettings.ChangeSetting", "PythonFolder", "D:\work\python"
```

### Using another Python

To use a Python of your own instead, such as a venv with pinned packages or your organisation's Python: install a [python.org](https://www.python.org/downloads/windows/) Python 3.11 or later, **of Excel's bitness**, make a venv and set `PythonVenv` to its folder:

```
py -3.14 -m venv %USERPROFILE%\asfpy-venv
%USERPROFILE%\asfpy-venv\Scripts\python -m pip install numpy pandas pywin32
```

```vb
Application.Run "ASF.xlam!ASFSettings.ChangeSetting", "PythonVenv", Environ$("USERPROFILE") & "\asfpy-venv"
```

For 32-bit Excel, use `py -3.14-32`. The Microsoft Store Python and free-threaded builds don't work. The package console then runs that venv's pip. Set `PythonVenv` back to empty to return to the add-in's own Python.

---

## Troubleshooting

When Python can't start, Excel shows the reason when it loads the add-in. `%TEMP%\asfpy.log` (type it in File Explorer's address bar) has the details: what was loaded, each file of the Python folder, and each error.

| You see | Why, and what to do |
|---|---|
| Excel says the file isn't a valid add-in, or nothing happens | The `.xll` doesn't match Excel's bitness, or it is still blocked: see steps 1 and 3 |
| "the Python folder gave no worksheet functions" | No file of the Python folder marks a function with `@xl_func`, or a file didn't import: the log says why. Download `.py` files as raw files: a page saved from the browser is HTML, and `name (1).py` or `name.py.txt` isn't the name you want |
| `#NAME?` in cells | The add-in isn't loaded, or the function isn't registered: see the log. For `PY_CALL`, the name isn't a marked function of the Python folder |
| `#VALUE!` | The Python function raised an exception: **Why this error?** on the ASF tab shows it |
| "another add-in has already loaded python3XX.dll" | Excel runs one Python per process. Another Python add-in (PyXLL, xlOil, xlwings in-process) is loaded: untick it |
| "another copy of this add-in already runs Python" | You loaded a second copy of the `.xll`: restart Excel and keep one |
| `ImportError` in the log | The package isn't installed in the add-in's Python: see [Packages](#6-packages) |

**Security.** The add-in runs the Python files of your Python folder with your permissions, as any program you install does. Only put files you trust there, and only install packages you trust. Python never imports from the folder of the workbook you opened, nor from the environment variables `PYTHONPATH` and `PYTHONSTARTUP`.

---

## LibreOffice

[`ASF.oxt`](Add-ins.html#libreoffice) carries Python support too, through LibreOffice's own Python (on Linux the system's, where numpy and pandas install as usual). It gives exported functions and `py`, not `@xl_func` worksheet functions. The Python folder is `%APPDATA%\ASF\python` on Windows and `~/.config/ASF/python` elsewhere:

```python
# stats.py in the Python folder
from asfpy import export

@export
def scale(values: list, factor: float):
    return [v * factor for v in values]
```

```vb
GlobalScope.BasicLibraries.LoadLibrary "ASF"
engine.InjectHostObject "py", LOpython.PyRoot(1)                   ' level 2 also imports modules
MsgBox engine.Run(engine.Compile("return py.call('stats.scale', [1, 2, 3], 10).length;"))   ' => 3
r = LOpython.PyCall("stats.scale", Array(1, 2), 10)                ' from Basic: rows from 1
MsgBox LOpython.PyReload()                                          ' after the modules changed
```

In a cell, `=PY_CALL("stats.scale"; A1:A3; 10)` calls the same function once the `ASF` library is loaded, for example by a macro on Tools > Customize > Events > Start Application. It takes up to 15 arguments. `LOpython.PyLastError()` says why a call failed.
