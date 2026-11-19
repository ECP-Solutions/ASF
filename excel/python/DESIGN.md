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

ASF calls host objects through `CallByName` (`ASF_Runtime.InvokeHost`, up to 30 arguments), so any `IDispatch` object works, including one that stands for Python. Four details of today's engine shape the design. They were checked against `src/` at 4.0.2:

| Behaviour | Where | Consequence for Python |
|---|---|---|
| With `AppAccess` off, an injected object becomes its type name (a string) | `ASF_Runtime.CastInjectedVar` | Injecting `py` alone does nothing. The engine needs a way to inject an object the host trusts while `AppAccess` stays off (3.3) |
| Method names on host objects are lower-cased before `CallByName` | `ASF_Runtime.TryCallHost` | COM ignores case, Python doesn't: `pd.DataFrame(...)` would arrive as `dataframe`. Pass the name as written (3.3) |
| `import` is a keyword even after a dot: `py.import(...)` does not compile | `ASF_Compiler` | The method gets another name, `py.module(...)` (3.3) |
| An error while reading a host property gives `undefined` | `ASF_Runtime.MemberValue` | A Python exception on an attribute read is silent. The proxy keeps the last exception for **Why this error?** |

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
  asfpy/      the Python package: decorators, conversions, object handles, reload
  vba/        ASFPython.bas: PY_CALL and the `py` object for ASF scripts
  tests/      Python tests (any OS), a fake Excel host for the .xll, export and size checks
```

`ASF.xlam` stays pure VBA and works without Python. `asfpy.xll` is a second, optional add-in. `ASF.xlam` uses it when it is loaded and says so when it is not.

### 3.1 `asfpy.xll`: the native part

**Finding and starting Python**

- **Stable ABI, `Py_LIMITED_API` 3.11.** One `.xll` then works with every Python from 3.11. 3.11 is the floor because it adds the buffer protocol (`PyObject_GetBuffer`) to the limited API, which numpy arrays need for fast conversion. 3.10 reached its end of life in October 2026.
- **Embedding is the weak spot of the stable ABI.** The limited API covers calling Python, not configuring it. `PyConfig` is not in the limited API, and its layout changes between versions. `Py_SetPythonHome` and `Py_SetProgramName` are deprecated. Two ways out:
  1. keep everything in the limited API except start-up. Read the version with `Py_GetVersion`, then call that version's `PyConfig` functions through `GetProcAddress`, with a small table of struct layouts per version; or
  2. start Python with defaults, then let a bootstrap written in Python fix `sys.path` and `sys.flags`-like settings before any user module is imported.

  Option 2 is simpler, but it can't set `safe_path` or `isolated` before the first import, and those flags matter for security (3.4). The first spike decides (phase 1).
- **The Python is set explicitly, never taken from `PATH`.** The setting is a venv folder. Its `pyvenv.cfg` names the base install (`home = …`), and that folder holds `python3.dll`. The DLL is loaded with its full path and `LoadLibraryExW(…, LOAD_WITH_ALTERED_SEARCH_PATH)`, so that `python3.dll` finds `python3XX.dll` next to it. `SetDllDirectory` is never called: it changes DLL loading for all of Excel.
- **Supported Pythons:** python.org installs and venvs made from them. Conda works if its `Library\bin` is added with `os.add_dll_directory`. **Not supported:**
  - the Microsoft Store Python, which installs into a protected app folder that another program's process can't rely on;
  - free-threaded builds (`3.13t` and later), which the stable ABI doesn't cover yet.
- **Same bitness as Excel.** A 64-bit Excel needs a 64-bit Python. The add-in checks the PE header of `python3.dll` before loading it and gives a clear message.
- **One Python per process.** Another add-in (PyXLL, xlOil) may already have started one. `asfpy.xll` checks for a loaded `python3*.dll` first. If one is there, it refuses and explains.
- **Python is started once and never finalized.** `Py_Finalize` with numpy or pandas loaded can crash, and they cannot start again in the same process. Reload means reloading modules (3.2), not restarting Python. The `.xll` pins itself in memory (`GetModuleHandleExW` with `GET_MODULE_HANDLE_EX_FLAG_PIN`). Unticking it in the Add-ins dialog then can't unload code that Python threads or VBA proxies still use.
- **Start-up cost.** Importing pandas takes about a second. `xlAutoOpen` imports the Python folder's modules on Excel's start, so Excel starts slower. Later, a manifest written on each load (names, arguments, descriptions) can register the functions at start and import each module on its first call.

**Worksheet functions**

- **A fixed pool of exported slots,** `asfpy_f0000`…`asfpy_f0999`, generated by a macro. Each one calls `dispatch(slot, args)`, and the Python side maps slots to functions. Excel-DNA has long used this approach. It needs no API hooking, which is easier to audit and does not alarm antivirus software. The pool size is a build constant.
- **The slots work because of the x64 calling convention.** The caller cleans the stack, so one C signature with the maximum number of arguments serves every function, and a slot reads only as many arguments as were registered. In 32-bit Excel (`__stdcall`) the callee cleans the stack, so each slot would need a fixed arity. **This design targets 64-bit Excel.** 32-bit is an open question (section 6).
- **Registration, with `xlfRegister`:**
  - every argument and the result are `Q` (`XLOPER12` by value). Ranges then arrive as values, as `ASFCells` expects. `U` (references) is only for functions that ask for them;
  - the type string's suffixes come from the decorator: `!` volatile, `$` thread-safe, `#` macro-equivalent. They can't be combined freely: `$` with `#` is refused, for example;
  - each argument description has at most 255 characters, as do the argument names together. `ASFRegister` already cuts them that way.
- **Memory:** results are returned with `xlbitDLLFree` and freed in `xlAutoFree12`. Thread-safe functions keep their result buffers per thread.
- **No Excel SDK files.** It declares `XLOPER12` itself and reaches `Excel12v` through `MdCallBack12`, the callback that Excel exports, as `xlcall.cpp` does. The structures are declared from Microsoft's documentation.

**The bridge to VBA**

- **One export, not three:** `AsfPy_Root(ByRef out As Variant) As Long`, which returns the root proxy object (`py`). The rest goes through `IDispatch`: `py.call(...)`, `py.module(...)`, `py.lastError`. The single `Declare` keeps two VBA traps out:
  - `Declare` converts `String` arguments to ANSI, which loses non-Latin text. A `Variant` passed `ByRef` arrives as a `VARIANT*`, untouched;
  - a `Declare` can't return a `Variant` reliably, so the result comes back through an argument.
- **Finding the DLL.** `Declare … Lib "asfpy.xll"` finds the module only if Excel has already loaded it. Otherwise it fails with error 53. `ASFPython.bas` catches that error and shows "the Python add-in is not loaded", not the VBA error.
- **A proxy `IDispatch` for Python objects:**
  - `GetIDsOfNames` maps to `getattr`, and `Invoke` to a call or attribute read. `DISPID_VALUE` is the default value, and `DISPID_NEWENUM` gives `For Each` over iterables;
  - `GetTypeInfoCount` returns 0;
  - VBA sends `DISPATCH_METHOD | DISPATCH_PROPERTYGET` together, so the proxy decides by what the attribute is: a callable is called, anything else is read;
  - names are matched by exact case first, then by case-insensitive match when it is unique. This covers VBA, whose editor changes the case of identifiers.
- **Lifetimes.** A proxy holds a reference to its Python object. `Release` takes the GIL before dropping it. Since Python is never finalized, a proxy VBA releases late, even after `xlAutoClose`, is still safe.

**Threads and the GIL**

- VBA, the ribbon and macros all run on Excel's main thread. Worksheet functions registered with `$` can run on calculation threads.
- After start-up, the add-in releases the GIL (`PyEval_SaveThread`). Every entry point takes it with `PyGILState_Ensure`.
- The GIL lets only one thread run Python at a time, so `thread_safe` brings little and stays off by default.
- Python must not call Excel's object model from a worksheet function. Excel ignores most such calls or blocks them. `xl_app()` works in macros only.

### 3.2 `asfpy`: the Python package

- `@xl_func`, `@xl_macro`, with type hints for the conversions and the docstring as the description in Insert Function. The names are common words, close to other add-ins' names, and the code is written from scratch. PyXLL's documentation and stub text are not copied.
- `@export`: marks a function that ASF scripts and `PY_CALL` may call (3.4).
- Conversions that follow ASF's own rules (`ASFCells`):

  | Cell / VBA | Python | Back to the cell |
  |---|---|---|
  | number | `float` (`int` with a type hint) | number; `int` beyond 2^53 loses digits |
  | date | `float` (the cell holds a serial number); `datetime` with a type hint | number |
  | text, boolean | `str`, `bool` | value; text over 32,767 characters is cut |
  | empty | `None` | `""` |
  | error value | not called: the error is the result, as in `ASF_RUN` | — |
  | one row or column | `list` | one column |
  | several rows and columns | `list` of rows (numpy/pandas with a type hint) | grid, spilled |
  | `nan`, `inf` | — | `#NUM!` |
  | other object | kept in the handle cache | `<DataFrame#12>`, a handle other functions accept |
  | exception | — | `#VALUE!`; the reason via **Why this error?** |

  From VBA (the `py` proxy), arrays arrive as `SAFEARRAY`s of `VARIANT`. ASF arrays have a lower bound of 1, ranges of 1, and VBA's `Array()` of 0. The proxy reads `LBound` and doesn't assume one. An ASF object (`ASF_Map`) is a VBA class with no type information, so `ASFPython.bas` converts it to a list of `[key, value]` pairs before the call, and Python gets a `dict`.
- **Handles are owned by their cell.** A handle is keyed by the calling cell (`xlfCaller`). When the cell is calculated again, its old object is dropped, so the cache doesn't grow with each recalculation. Handles are also scoped to the workbook.
- **Reload** re-imports the Python folder's modules and registers the functions again. Python itself stays (3.1).
- **Settings live where ASF's do.** They go under the same registry key as `SaveSetting` (`HKCU\Software\VB and VBA Program Settings\ASF`): the Python folder, the venv, and Python access. The `.xll` reads that key with `RegGetValueW`. An administrator then has one place to set, or lock, both add-ins. There is no `asfpy.cfg`.
- **Pure Python where possible:** tested on Linux in CI, with a fake Excel in place of the `.xll`.

### 3.3 ASF integration

**Worksheet functions:**

```
=PY_CALL("stats.zscore", A2:A500)            stats.py in the Python folder, exported function zscore
=forecast(B2:B40, 12)                        a function registered with @xl_func
=ASF_RUN("return py.call('stats.p90', $1);", A2:A500)
```

**ASF scripts** get a `py` object (when allowed, 3.4):

```js
// a formula script (ASF_RUN): exported functions only
return py.call('stats.zscore', $1);

// a library script with the header  // @asf python: modules
let pd = py.module('pandas');
let df = pd.DataFrame($1);
return df.describe().values.tolist();
```

**Changes to the ASF engine,** small and useful beyond Python:

1. **Trusted injection.** A host method such as `InjectHostObject(name, obj)` injects an object that stays an object while `AppAccess` is off. Only VBA can call it, not scripts, so the sandbox still decides what a script can reach. The add-in uses it for `py`.
2. **The method name as written.** `TryCallHost` passes the name with its case to `CallByName`. COM objects ignore case, so nothing else changes.
3. Optionally, **keywords after a dot** (`obj.import`, `obj.default`), as JavaScript allows. The design uses `py.module` and doesn't depend on this change.

The proxy never hands a raw COM object to ASF. A Python function that returns Excel's `Application` gives a script a `#VALUE!`, not Excel access. The ribbon gets **Python folder** and **Reload Python**. **Why this error?** also shows the Python traceback. The descriptions go into `ASFRegister`, as they do now.

### 3.4 Security: trust follows where the code came from

ASF can sandbox a script. **Python cannot be sandboxed:** any Python code can read files, open sockets and start programs. So the design sorts code by where it comes from, not by what it calls:

| Comes from | Trusted? | Gets |
|---|---|---|
| The workbook: `ASF_RUN` script text and cell values | No: anyone can send a workbook, and formulas run without a macro warning | `py.call` of exported functions only. `py.module` is undefined |
| The library folder: `.vas` scripts the user installed | Yes, like any installed script | With the header `// @asf python: modules`: `py.module` too |
| The Python folder: `.py` modules the user installed | Yes | Everything Python can do |

**Why there is no package allow-list for formulas.** The first draft let `py.import` reach `numpy`, `pandas` and `scipy` from any formula. A list of packages doesn't hold:
- `numpy.load(f, allow_pickle=True)` and `pandas.read_pickle` run code from a file;
- `pandas.read_csv` reads any path or URL, and `DataFrame.to_csv` writes one;
- `numpy.ctypeslib.load_library` loads a DLL;
- modules reach other modules through their attributes.

Making that safe would mean reviewing every callable of every package, version by version. The table above avoids the problem instead: the workbook can only pick which installed function runs, and with what data.

**Rules:**

1. **No Python source in workbooks.** There is no `PY_EVAL("…")` and no Python in cells. Excel's own `=PY()` is different: it runs in Microsoft's cloud sandbox.
2. **`PY_CALL` looks names up; it doesn't import them.** `"stats.zscore"` is found in the registry built when the Python folder loaded. A cell string never reaches `importlib`, so `PY_CALL("os.system", …)` finds nothing.
3. **Exported functions receive untrusted data.** The documentation gives authors a checklist: no paths, URLs, `eval` or pickle taken from arguments, and a size limit on the data.
4. **The import path is pinned.** `sys.path` holds the Python folder, the venv's `site-packages` and the standard library. It never holds `''` or the current folder. Excel's current folder is often the folder of the workbook just opened, such as Downloads, and a `numpy.py` placed next to a workbook must not be imported. `safe_path` and ignoring the environment (`PYTHONPATH`, `PYTHONSTARTUP`) follow from the start-up option chosen in 3.1. DLLs are loaded by full path only (3.1).
5. **Proxies stay inside the rules.** Attributes that start with `_` are refused. Module objects are not handed out except through `py.module`, and raw COM objects never. The origin table is the real barrier; this is defence in depth.
6. **Handles can't cross workbooks** (3.2). A handle typed into a cell names nothing outside its own workbook.
7. **Distribution.** Excel has blocked `.xll` files from the internet since 2023 (Mark of the Web), and signing doesn't change that. Users install through a script that unblocks the file, or from a trusted location. CI publishes the `.xll` with a checksum, and organisations can sign it for the "require signed add-ins" policy. Expect antivirus software to look hard at any unsigned `.xll` that loads Python: XLL files are a known malware vehicle.

### 3.5 Platforms

- **Excel for Windows, 64-bit:** the target, Excel 2016 to Microsoft 365.
- **Excel for Windows, 32-bit:** many Excel 2016 installs are 32-bit. It needs a 32-bit Python, a second build, and slots with fixed arities (3.1). Open question.
- **Excel for Mac:** an `.xll` cannot load there. Python would have to run in a separate process that VBA calls (as xlwings does), which is slower and a separate project. Mac keeps ASF without Python.
- **LibreOffice:**
  - **`PY_CALL`** can be a Basic function in the `Standard` library, as the `ASF.oxt` functions already are. It calls Python through the script provider (`getScript("vnd.sun.star.script:stats.py$zscore?language=Python&location=user")`). It must never use `location=document`: LibreOffice can embed Python macros in a document, and those come from the workbook.
  - **`py` in ASF scripts** needs a Python UNO object that implements `XInvocation`, and Basic's `CallByName` working on it. This needs a spike.
  - **`@xl_func` functions:** Calc gets native functions from Python only through a Calc Add-In, a UNO component whose interface is compiled from IDL. That interface is fixed at build time, so functions can't be added at run time. Out of scope.
  - **Packages:** on Linux, LibreOffice uses the system Python, so numpy and pandas install as usual. On Windows and macOS it ships its own Python without pip, which makes packages awkward. LibreOffice is a cheap target on Linux only.

## 4. Build or reuse

| Option | Licence | Fit |
|---|---|---|
| **Own small `.xll` (proposed)** | MIT | Fits ASF's zero-dependency style. Covers exactly what ASF needs. The stable ABI gives one binary for all Pythons from 3.11 |
| pywin32 COM server | PSF-like | A Python class registered as a COM server gives VBA an `IDispatch` with no C at all. Run as a separate process, it isolates crashes and doesn't have to match Excel's bitness; each call crosses processes. Needs pywin32 and a per-user COM registration. No native worksheet functions, so `PY_CALL` only |
| [xlOil](https://github.com/cunnane/xloil) | Listed as Apache-2.0 on PyPI; check its `LICENSE` | An open source in-process Python add-in, close to PyXLL. Large C++ codebase. ASF would depend on its release cycle |
| xlwings (open source core) | BSD-3 | Python runs out of process over COM, with a VBA module per workbook. Slower; worksheet functions need VBA in the workbook |
| Excel-DNA + Python.NET | zlib / MIT | Adds .NET as a dependency |

Recommendation: write our own small add-in, and read xlOil and Excel-DNA for ideas, not code. **Hide the transport behind the `py` object.** VBA and ASF see only an `IDispatch`, so most of the work doesn't depend on the `.xll`: the Python package, conversions, the security rules, `ASFPython.bas`, the engine changes and the tests. A pywin32 server can stand in for the `.xll` while that work is built and tried by users. The `.xll` then replaces it for speed and for native functions.

## 5. Phases

| Phase | Delivers | Done when |
|---|---|---|
| 0 | Investigation, design, toolchain spike | This page; the MinGW build |
| 1 | **Excel spike.** `asfpy.xll` loads a venv's Python by full path, registers one slot function, and gives VBA the root proxy through `AsfPy_Root` | In Excel 2016 x64 and Microsoft 365: a formula returns a value from Python; VBA calls `py.call`; the start-up option of 3.1 is chosen; the cost of one call and of start-up is measured; Excel closes cleanly with numpy loaded |
| 2 | `PY_CALL`; `@xl_func` and `@export` with plain values and arrays; errors and **Why this error?**; reload; registry settings | Python tests on Linux; the `.xll` driven by a fake Excel host in CI (see below); a checklist run in Excel |
| 3 | Engine changes (3.3); `py` in ASF scripts; the origin rules and the `python` header; ribbon buttons | Plus ASF test cases with a fake `py` object in LibreOffice and Excel |
| 4 | numpy and pandas conversions, object handles, `thread_safe`, async functions, `@xl_macro` | |
| 5 | LibreOffice `PY_CALL` and `py` (after its spike); later, if ever, the ribbon from Python, RTD and task panes | |

Phases 2 and 3 already give users what matters most: Python functions in formulas and in ASF scripts.

**Testing the `.xll` without Excel.** `xlcall.cpp` finds `MdCallBack12` with `GetProcAddress(GetModuleHandle(NULL), …)`, that is, in the program that loaded the add-in. A small test program can export its own `MdCallBack12`, answer `xlfRegister` and the other calls the add-in makes, load `asfpy.xll` and call its slots. Built with MinGW, it can run in CI under Wine with the Windows embeddable Python. This is to be confirmed in phase 2. The GitHub-hosted Windows runners have no Excel, so the last check stays a self-hosted runner or a checklist run by hand before each release.

## 6. Open questions

- **32-bit Excel:** support it, with a second build and fixed-arity slots, or require 64-bit Excel for Python?
- **The start-up option** (3.1): per-version `PyConfig` through `GetProcAddress`, or a Python bootstrap. Phase 1 decides.
- **The name:** `asfpy.xll` beside `ASF.xlam` (proposed, as `ASF.xlam` must keep working without Python), or one `ASF.xll` one day.
- **`py.module` in library scripts:** is the header enough, or should an administrator be able to turn it off with a setting?
- **pywin32 first?** Whether to build phases 2 and 3 on a pywin32 server before the `.xll` (section 4).
