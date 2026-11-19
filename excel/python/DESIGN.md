# Python for ASF in Excel: investigation and design

Status: proposal. Nothing here is built yet.

Goal: Excel users call Python modules from their worksheets and from ASF scripts, through an open source add-in (MIT, like the rest of ASF). PyXLL, a commercial add-in, is the reference for what is possible. This page explains how PyXLL works and then proposes how ASF can do the same.

## 1. How PyXLL works

### What was looked at

The PyXLL 5.13.0 download for 64-bit Excel and Python 3.14:

| File | What it is |
|---|---|
| `pyxll.xll` | The add-in: a native 64-bit Windows DLL (PE32+), written in C++ |
| `pyxll.cfg` | The settings: which Python, which modules, the ribbon, logging, the license key |
| `pyxll-5.13.0-cp314-none-any.whl` | The `pyxll` Python package. It contains **stubs only**, so code can be imported and tested outside Excel. The real module is compiled into the `.xll` |
| `examples/`, `examples.xlsx` | Examples of every feature |
| `readme.pdf`, `license-agreement.pdf` | Quick start, open source notices, license |

The license forbids reverse engineering, decompiling and disassembling the software (clause 3). So this investigation uses only the documentation, the settings file, the public stub module, the examples, and the **export and import tables** of the DLL, which Windows itself reads to load it. No code was disassembled, and nothing below comes from PyXLL's code. Where a point is an inference, it says so.

### The layers

```
Excel.exe
 └─ pyxll.xll                 native add-in, loaded in Excel's process
     ├─ Excel C API           xlAutoOpen, xlfRegister, Excel12v: worksheet functions, macros
     ├─ COM server            ribbon, task panes, RTD (DllGetClassObject, DllRegisterServer)
     └─ python3XX.dll         loaded at run time, from the Python set in pyxll.cfg
         └─ module "pyxll"    built into the .xll: decorators, conversions, xl_app()...
             └─ the user's modules, listed in pyxll.cfg
```

**1. A native add-in in Excel's process.** An `.xll` is a DLL that Excel loads and calls through the Excel C API. The export table has the standard entry points: `xlAutoOpen`, `xlAutoClose`, `xlAutoFree12`, `xlAddInManagerInfo`, `xlAutoRemove`, `SetExcel12EntryPt`. It imports `XLCall32.dll`, Excel's C API library. Because Python runs **inside** Excel, a call costs no process switch or COM round trip. This is PyXLL's main difference from xlwings, which drives Excel from a separate Python process. Excel, the add-in and Python must all have the same bitness.

**2. Python is loaded at run time.** The import table has no `python3XX.dll`. The add-in finds the runtime from the `executable` setting (or `pythonhome` and `dll`), loads it and starts it. The wheel is tagged `cp314` and the download is per Python version. So the add-in is compiled against the full Python C API of one Python version, not against the stable ABI. It sets `sys.flags`: `ignore_environment` keeps a global `PYTHONPATH` out, and `inspect` keeps a `SystemExit` from closing Excel.

**3. The `pyxll` module is built into the add-in.** It is a built-in module created by the DLL, so it is always importable inside Excel. The wheel is a copy with empty bodies, for editors, type checkers and unit tests.

**4. Registration by decorators.** At start, the add-in imports the modules listed in `modules =`. Decorators record the functions:

| Decorator | Gives Excel |
|---|---|
| `@xl_func` | a worksheet function. Options: signature or type hints, category, argument descriptions, `volatile`, `thread_safe`, `macro`, `auto_resize`, `transpose`, LRU cache, formatter, how NaN, inf and errors become cell values |
| `@xl_macro` | a macro (`Application.Run`, buttons) |
| `@xl_menu` | a menu item under Add-ins |
| `@xl_on_open`, `@xl_on_close`, `@xl_on_reload` | events of the add-in |
| `@xl_arg_type`, `@xl_return_type` | custom conversions between cells and Python types |

Each function is then registered with Excel through `xlfRegister`. Excel needs an **exported procedure name** for each registered function. The DLL exports only 21 names, so the entry points for the user's functions are supplied at run time. The open source notices include MinHook, an API hooking library, which suggests the add-in intercepts how Excel resolves procedure names. This is an inference and was not verified. `MacroWrapper` and `PlotOnAction` look like generic entry points for macros and chart actions.

**5. Conversions.** Arguments arrive as `XLOPER12` values and are converted by type: `float`, `int`, `str`, `bool`, `var`, `float[]` and 2D arrays, dates, `numpy` arrays, pandas `DataFrame` and `Series`, and custom types. A Python object that has no cell value is kept in an **object cache** and the cell shows a handle, which other functions take as an argument. That lets formulas pass whole objects (a model, a DataFrame) from cell to cell. `nan_value`, `none_value` and `propagate_errors` decide how special values come back.

**6. Concurrency.** `thread_safe=True` registers the function for multi-threaded recalculation. Async functions use Excel's asynchronous UDFs (`xlAsyncReturn`). The export `xlCalculationCanceled` handles Excel's cancel event, and `ProcessAsyncCalls` runs work queued for Excel's main thread (`schedule_call`, `async_call`). Python's GIL serializes Python code itself.

**7. Talking to Excel.** `xl_app()` returns Excel's `Application` object for `win32com` or `comtypes`. The import of `OLEACC.dll` suggests it gets the object from Excel's window with `AccessibleObjectFromWindow`, the usual way to do this inside Excel (inference). Macros can then write cells, format and add sheets, as VBA does.

**8. COM server.** `DllGetClassObject` and `DllRegisterServer` make the DLL a COM server too. It needs that for what Excel only offers to COM add-ins: the **ribbon** (from the XML in `ribbon =`), **custom task panes** (Qt, wx or Tk UIs), **ActiveX controls**, and **RTD** real-time data functions.

**9. Development comforts.** `reload()` and `auto_reload` reload the modules without restarting Excel, with `deep_reload` for their dependencies. `get_last_error(cell)` returns the exception behind an error cell. `error_handler` decides what the cell shows. Logging goes to a folder.

**10. Typing tips.** The export `IntellisenseRegistrationInfo` and the RegistrationInfoGuid notice show it feeds the Excel-DNA IntelliSense add-in. ASF already supports that add-in, from the `_IntelliSense_` sheet.

**11. Deployment.** `external_config` (paths or URLs) keeps settings central. `startup_script` (`.bat`, `.ps1`) installs packages before Python starts. There is a license key.

**Limits.** Windows only (an `.xll` is a Windows DLL). One Python per Excel process. Bitness must match. One build per Python version.

## 2. What ASF has, and what is missing

ASF runs entirely in VBA. Its security rests on one idea: **a script is data**. A workbook's formulas run with no macro warning, so every script is sandboxed. It has no Office access (`AppAccess` off), imports only from the library folder, and stops at a step limit.

ASF calls host objects through `CallByName`, which means any `IDispatch` object (`ASF_Runtime.InvokeHost`). An ASF script can therefore use an object that represents Python, as long as the object is a COM object.

Missing:
- a way to run Python in Excel's process;
- a way for VBA, and so for ASF, to call it;
- Python functions as Excel worksheet functions;
- a security model for code that cannot be sandboxed.

## 3. Proposal: `asfpy`

Three parts, all MIT, in `excel/python/`:

```
excel/python/
  xll/        asfpy.xll: C, about 2–3k lines, no Excel SDK files, built with MinGW or MSVC
  asfpy/      the Python package: decorators, conversions, object handles, settings, reload
  vba/        ASFPython.bas: PY_CALL and the `py` object for ASF scripts
  tests/      Python tests (any OS), export and size checks of the built .xll
```

### 3.1 `asfpy.xll`: the native part

- **Loads Python through the stable ABI** (`python3.dll`, `Py_LIMITED_API` 3.10). One `.xll` then works with every Python from 3.10, where PyXLL ships a build per version. Python is found from `asfpy.cfg` (an `executable`, or a venv folder), then `PATH`.
- **A fixed pool of exported slots** for worksheet functions: `asfpy_f0000`…`asfpy_f0999`, generated by a macro. Each one calls `dispatch(slot, args)`, and the Python side maps slots to functions. This is the approach Excel-DNA has long used. It needs no API hooking, which is easier to audit and does not alarm antivirus software. 1000 functions is plenty; the pool size is a build constant.
- **No Excel SDK files.** It declares `XLOPER12` itself and reaches `Excel12v` through `MdCallBack12`, the callback that Excel exports, as `xlcall.cpp` does.
- **Exports for VBA:** `AsfPy_Call(name, args) As Variant`, `AsfPy_Import(module) As Object`, `AsfPy_LastError() As String`. VBA declares them with `Declare PtrSafe Function … Lib "asfpy.xll"`. Windows resolves the bare name to the module Excel already loaded.
- **A proxy `IDispatch` for Python objects.** `GetIDsOfNames` maps to `getattr`, and `Invoke` to a call or attribute read. Through it, ASF's `CallByName` reaches any Python object with no change to the ASF engine.
- Standard exports: `xlAutoOpen`, `xlAutoClose`, `xlAutoFree12`, `xlAddInManagerInfo`.

**Spike, done.** A minimal `.xll` with 200 exported slots, `xlAutoOpen` and run-time loading of `python3.dll` cross-compiles on Linux with `x86_64-w64-mingw32-gcc`. The result imports only `KERNEL32` and `msvcrt`. So the add-in can be built in GitHub Actions on Linux. Running it still needs Windows and Excel.

### 3.2 `asfpy`: the Python package

- `@xl_func`, `@xl_macro`, with type hints for the conversions and the docstring as the description in Insert Function. The names follow PyXLL's, so existing code ports with an import change. The code is written from scratch.
- `@export`: marks a function that ASF scripts and `PY_CALL` may call (see Security).
- Conversions that follow ASF's own rules (`ASFCells`):

  | Cell / VBA | Python | Back to the cell |
  |---|---|---|
  | number | `float` (`int` with a type hint) | number |
  | text, boolean | `str`, `bool` | value |
  | empty | `None` | `""` |
  | error value | not called: the error is the result, as in `ASF_RUN` | — |
  | one row or column | `list` | one column |
  | several rows and columns | `list` of rows (numpy/pandas with a type hint) | grid, spilled |
  | other object | kept in the handle cache | `<DataFrame#12>`, a handle other functions accept |
  | exception | — | `#VALUE!`; the reason via `PY_LASTERROR` or the ribbon |

- Reload without restarting Excel, settings and logging.
- Pure Python: tested on Linux in CI, with a fake Excel in place of the `.xll`.

### 3.3 ASF integration

**Worksheet functions:**

```
=PY_CALL("stats.zscore", A2:A500)            stats.py in the Python folder, function zscore
=forecast(B2:B40, 12)                        a function registered with @xl_func
=ASF_RUN("let np = py.import('numpy'); return np.percentile($1, 90);", A2:A500)
```

**ASF scripts** get a `py` object (when allowed):

```js
let pd = py.import('pandas');
let df = pd.DataFrame($1);
return df.describe().values.tolist();
```

This works through the proxy `IDispatch` and ASF's existing `CallByName`. The ASF engine needs one new option, `PythonAccess`, set by the add-in, that injects `py`. The ribbon gets **Python folder**, **Reload Python** and **Why this error?** (which also covers Python errors). The descriptions go into `ASFRegister` as they do now.

### 3.4 Security: the decisive difference

ASF can sandbox a script. **Python cannot be sandboxed:** any Python code can read files, open sockets and start programs. So the rule that a received workbook never runs code the user did not install must hold by construction:

1. **No Python source in workbooks.** There is no `PY_EVAL("…")` and no Python in cells, unlike Excel's own `=PY()`, which runs in Microsoft's cloud sandbox.
2. **Only installed functions.** `PY_CALL` and `py` reach only modules of the Python folder (`%APPDATA%\ASF\python`, a setting an administrator can fix), and only functions marked `@export` or `@xl_func`. `py.import` is limited to an allow-list of packages (`numpy`, `pandas`, `scipy`, `statistics`… by default), because `py.import('os')` would undo the sandbox.
3. **Off by default for `ASF_RUN`.** The `PythonAccess` setting is `0`, so `py` is undefined in formula scripts. Library scripts may ask for it with a header line, `// @asf python: true`, the way `appAccess` works now.
4. **Proxies stay inside the rules.** Attributes that start with `_` are refused through the proxy, which closes `__class__.__subclasses__()` and similar escapes. The allow-list is the real barrier; this is defence in depth.
5. **Signed builds.** CI publishes the `.xll` with a checksum, and organisations can sign it.

### 3.5 Platforms

- **Excel for Windows:** the target.
- **Excel for Mac:** an `.xll` cannot load there. Python would have to run in a separate process that VBA calls (as xlwings does), which is slower and a separate project. Mac keeps ASF without Python.
- **LibreOffice:** it ships Python and can call Python scripts from Basic through its script provider. `ASF.oxt` could get the same `py` object at little cost. This is a cheap second target.

## 4. Build or reuse

| Option | Licence | Fit |
|---|---|---|
| **Own small `.xll` (proposed)** | MIT | Fits ASF's zero-dependency style. Covers exactly what ASF needs. Stable ABI gives one binary for all Pythons |
| [xlOil](https://github.com/cunnane/xloil) | Listed as Apache-2.0 on PyPI; check its `LICENSE` | An open source in-process Python add-in, close to PyXLL. Large C++ codebase. ASF would depend on its release cycle |
| xlwings (open source core) | BSD-3 | Python runs out of process over COM, with a VBA module per workbook. Slower; worksheet functions need VBA in the workbook |
| Excel-DNA + Python.NET | zlib / MIT | Adds .NET as a dependency |

Recommendation: write our own small add-in, and read xlOil and Excel-DNA for ideas, not code.

## 5. Phases

| Phase | Delivers | Checked by |
|---|---|---|
| 0 | Investigation, design, toolchain spike | This page; the MinGW build |
| 1 | `asfpy.xll` loads Python; `PY_CALL`; `@xl_func` with plain values and arrays; errors; reload; `asfpy.cfg` | Python tests on Linux; Linux CI build; manual test in Excel for Windows |
| 2 | Proxy `IDispatch`; `py` in ASF scripts; `PythonAccess` and the header; allow-list; ribbon buttons | Plus ASF test cases with a fake `py` |
| 3 | numpy and pandas conversions, object handles, `thread_safe`, async functions, `@xl_macro` | |
| 4 | Ribbon from Python, RTD, task panes (COM server), LibreOffice `py` | |

Phases 1 and 2 already give users what matters most: Python libraries in formulas and in ASF scripts. PyXLL's own extras (task panes, ActiveX, plotting) can come later, if ever.

## 6. Open questions

- The minimum Python version: 3.10 for the stable ABI, or later?
- The name: `asfpy.xll`, or one `ASF.xll` that would one day carry the whole add-in?
- Which packages the default allow-list should hold.
- Testing in Excel: the GitHub-hosted Windows runners have no Excel, so we need a self-hosted runner or a release checklist run by hand.
