# Python for ASF in Excel: investigation and design

Status: phases 1a to 3 are built (`xll/`, `ASF.xlam`, checks in `tests/`, the Excel checklist in `tests/excel/README.md`) and pass in 64-bit Excel 16.0; their 32-bit and Excel 2016 runs, and CI, are still to do. Phase 4 is built, passes the fake host in 64-bit and 32-bit, and passes its checklist in 64-bit Excel 16.0. Phase 5 (LibreOffice) is built: `py` in ASF scripts and `PY_CALL` pass in a headless LibreOffice, Calc cells are still to check by hand. The ribbon from Python, RTD and task panes stay out.

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
| Host calls go through `CallByName` | `ASF_Runtime.InvokeHost` | In Excel, `CallByName` reports a failed call as error 440 "Automation error", although it gets the message in `EXCEPINFO`; a late-bound call gets it, as error vbObjectError + 1001 (phase 1b's checklist). So the text always comes from `py.lastError` |

Missing:
- a way to run Python in Excel's process;
- a way for VBA, and so for ASF, to call it;
- Python functions as Excel worksheet functions;
- a security model for code that cannot be sandboxed.

## 3. Proposal: `asfpy`

Three parts, all MIT, in `excel/python/`:

```
excel/python/
  xll/        asfpy.xll: C, about 2–3k lines, no Excel SDK files, built with MinGW or MSVC; PY_CALL
  asfpy/      the Python package: decorators, conversions, object handles, reload
  vba/        ASFPython.bas: the `py` object for ASF scripts
  tests/      Python tests (any OS), a fake Excel host for the .xll, export and size checks
```

Since phase 2, `PY_CALL` is registered by `asfpy.xll` itself, not written in VBA: it works whenever the `.xll` is loaded, without `ASF.xlam`, and a call doesn't pass through VBA. Until the package needs files of its own (object handles, stubs for editors), the Python side lives in `xll/boot.py`, which `build.py` compiles into the add-in and `tests/unit/test_boot.py` tests on any OS.

`ASF.xlam` stays pure VBA and works without Python. `asfpy.xll` is a second, optional add-in, built for 64-bit and 32-bit Excel (3.1). `ASF.xlam` uses it when it is loaded and says so when it is not.

### 3.1 `asfpy.xll`: the native part

**Finding and starting Python**

- **Stable ABI, `Py_LIMITED_API` 3.11.** One `.xll` then works with every Python from 3.11. 3.11 is the floor because it adds the buffer protocol (`PyObject_GetBuffer`) to the limited API, which numpy arrays need for fast conversion. 3.10 reached its end of life in October 2026.
- **Embedding is the weak spot of the stable ABI.** The limited API covers calling Python, not configuring it. `PyConfig` is not in the limited API, and its layout changes between versions. `Py_SetPythonHome` and `Py_SetProgramName` are deprecated. Two ways out:
  1. keep everything in the limited API except start-up. Read the version with `Py_GetVersion`, then call that version's `PyConfig` functions through `GetProcAddress`, with a small table of struct layouts per version; or
  2. start Python with defaults, then let a bootstrap written in Python fix `sys.path` and `sys.flags`-like settings before any user module is imported.

  Option 2 is simpler, but it can't set `safe_path` or `isolated` before the first import, and those flags matter for security (3.4).

  **Chosen in phase 1a: neither, a better third way.** Python 3.14 added a configuration API that works by name (`PyInitConfig_SetInt(cfg, "safe_path", 1)`, PEP 741), so no struct layouts are needed. For 3.11 to 3.13, which only get security fixes now, the add-in sets the global flags those versions export (`Py_IsolatedFlag`, `Py_IgnoreEnvironmentFlag`, `Py_NoSiteFlag`, `Py_NoUserSiteDirectory`) and calls `Py_SetPythonHome` before `Py_InitializeEx(0)`. Their exports won't change any more. Either way Python is isolated from the environment and doesn't import `site`; the bootstrap adds the Python's `site-packages` with `site.addsitedir`. An embedded Python adds no script folder to `sys.path`, and the bootstrap removes `''` and the current folder in case. Checked with 3.11, 3.12, 3.14 and 3.15.
- **The Python is the add-in's own, or set explicitly, never taken from `PATH`.** By default it is the Python the add-in carries (below). The setting `PythonVenv` names another: a venv folder, or a Python install. A venv's `pyvenv.cfg` names the base install (`home = …`), and that folder holds `python3NN.dll`. The add-in loads that DLL directly, with its full path and `LoadLibraryExW(…, LOAD_WITH_ALTERED_SEARCH_PATH)`, and finds every function with `GetProcAddress`. It doesn't go through `python3.dll`: its exports forward to `python3NN.dll`, and Windows would look for that DLL along Excel's search path, not next to `python3.dll`. The functions used are all stable-ABI functions, which `python3NN.dll` exports under the same names. A free-threaded `python3NNt.dll`, and the `python3t.dll` that 3.15 ships, are skipped. `SetDllDirectory` is never called: it changes DLL loading for all of Excel.
- **Supported Pythons:** python.org installs and venvs made from them. Conda works if its `Library\bin` is added with `os.add_dll_directory`. **Not supported:**
  - the Microsoft Store Python, which installs into a protected app folder that another program's process can't rely on;
  - free-threaded builds (`3.13t` and later), which the stable ABI doesn't cover yet.
- **Same bitness as Excel.** There are two builds: `asfpy64.xll` for 64-bit Excel with a 64-bit Python, and `asfpy32.xll` for 32-bit Excel with a 32-bit Python. The add-in checks the PE header of `python3.dll` before loading it and gives a clear message. In 32-bit Excel, Python shares Excel's address space of at most 4 GB, so large DataFrames run out of memory sooner.
- **Installing: the add-in carries its Python.** Nobody has to install Python or make a venv. The idea comes from [sancarn/xll-python](https://github.com/sancarn/xll-python), whose single `.xll` holds CPython, unpacks it under `%LOCALAPPDATA%` on first use, and runs pip in a separate process. Its code is under a source-available licence with conditions of its own, and none of it is copied: it is a .NET (Excel-DNA) add-in for 64-bit Excel, while `asfpy.xll` stays plain C for both bitnesses. What `asfpy.xll` does:
  - **The payload.** `build.py` takes the CPython that `xll/runtime.json` pins (3.14.10 now): python.org's own Windows builds as published on nuget.org, the packages `python` and `pythonx86`, checked against their SHA-256. They are complete installs with pip already in `site-packages`, so nothing is extracted from MSIs or bootstrapped. C headers, import libraries and `ensurepip` are left out. The rest (about 1,100 files, 40 MB) goes into a cabinet (MSZIP) that `build.py` writes itself, stored as the resource `ASFPY_RUNTIME` (`windres` for MinGW, `rc` for MSVC). The add-ins grow to 13.9 MB (64-bit) and 12 MB (32-bit). `build.py --no-runtime` builds them without it, for `PythonVenv` only.
  - **First load.** `xlAutoOpen` unpacks the resource with `cabinet.dll` (FDI, part of every Windows: no unzip code of its own and no other process) into `%LOCALAPPDATA%\ASF\runtime\python-3.14.10-64-<hash>`. It writes into a temporary folder and renames it when complete, under a named mutex for two Excels starting at once, so a folder of that name is always whole. That took 430 ms under Wine. Later loads only check that the folder exists. The name changes only when the Python does; old folders stay, with their packages, until removed by hand.
  - **Packages.** The command `AsfPyPackages` (**Python packages** on the ASF tab) opens a console window that runs pip for the add-in's Python (`python.exe -I`, a separate process), so downloads never block Excel. Packages go to that Python's own `Lib\site-packages`, which the bootstrap adds. Two limits: Windows locks a package's DLLs while Excel has it loaded, so upgrading or removing such a package needs Excel closed; and a later add-in that carries another Python starts with no packages (`pip freeze` in the old one lists them). Offline or behind a proxy, pip takes wheel files and `--index-url`, as usual. With `PythonVenv` set, the command runs that Python's `python.exe`.
  - **`PythonVenv` still wins** when it is set: an organisation's own Python, a conda environment, or a venv with pinned packages.
  - **Trust doesn't change.** The runtime folder is writable by the user, as a venv is: whoever can write there can already run code as that user. Installing a package remains the user's decision.
  - Not taken from xll-python: its WebView2 package browser, IDE, console, self-updater, and its compiler that packs a project with its packages into an `.xll` of its own. The last is worth keeping in mind for organisations that deploy a fixed set of packages.
- **One Python per process.** Another add-in (PyXLL, xlOil) may already have started one. `asfpy.xll` checks for a loaded `python3*.dll` first. If one is there, it refuses and explains. When the one there was loaded by another copy of `asfpy.xll` (a module that exports `AsfPy_Root`), the message names that copy and says to restart Excel: browsing to a second copy is the usual way to get there.
- **Python is started once and never finalized.** `Py_Finalize` with numpy or pandas loaded can crash, and they cannot start again in the same process. Reload means reloading modules (3.2), not restarting Python. The `.xll` pins itself in memory (`GetModuleHandleExW` with `GET_MODULE_HANDLE_EX_FLAG_PIN`). Unticking it in the Add-ins dialog then can't unload code that Python threads or VBA proxies still use. `xlAutoClose` unregisters the functions (`xlfUnregister`), so Excel no longer calls them, and their cells show `#VALUE!` while the add-in is off. Excel keeps a function's name for the session: deleting it (`xlfSetName`), even after registering the function again as hidden as Excel-DNA does, still gave `#VALUE!` in Excel, not `#NAME?`. `xlAutoOpen` registers the functions again when the add-in is ticked.
- **Start-up cost.** Importing pandas takes about a second. `xlAutoOpen` imports the Python folder's modules on Excel's start, so Excel starts slower. Later, a manifest written on each load (names, arguments, descriptions) can register the functions at start and import each module on its first call.

**Worksheet functions**

- **A fixed pool of exported slots,** `asfpy_f0000`…`asfpy_f0999`, generated by `build.py`. Each one calls `dispatch(slot, args)`, and the Python side maps slots to functions. Excel-DNA has long used this approach. It needs no API hooking, which is easier to audit and does not alarm antivirus software. The pool size is a build constant. When the Python folder gives more functions than slots, the add-in registers the first ones and says how many it left out, in an alert and in the log.
- **Slots in each bitness.** One slot serves functions with any number of arguments, but the two calling conventions make that work differently:
  - **64-bit:** the caller cleans the stack. One C signature with the maximum number of arguments serves every function, and a slot reads only as many arguments as were registered. Plain C.
  - **32-bit:** Excel calls with `__stdcall`, so the function itself removes its arguments, 4 bytes for each `XLOPER12` pointer. A slot must remove exactly as many as were registered. Each slot is a two-instruction stub in assembly (`mov eax, slot` then `jmp common`). The common entry calls `dispatch(slot, args)`, reads the slot's argument count from a table filled at registration, and returns past exactly that many arguments. That is about 30 lines of GNU assembler, built with `i686-w64-mingw32`; MSVC builds the same instructions as naked functions with inline assembly. If the stub proves fragile, the fallback is plain C: the pool is split by argument count (slots for 0, 1, 2… arguments), and each function takes a slot of its own count. That wastes slots but needs no assembly.
  - **Export names:** `__stdcall` names are decorated (`_asfpy_f0000@8`). MinGW exports the plain names that Excel and VBA's `Declare` look for with `--kill-at`; MSVC with a `.def` file.
  - **Struct layouts:** `XLOPER12` and its members have different offsets in each bitness. Since the add-in declares them itself, `xloper.h` checks their sizes and offsets with `_Static_assert` in both builds: 32 bytes in each, the type at offset 24, the members after a pointer at 8 or 4.
- **Registration, with `xlfRegister`:**
  - every argument and the result are `Q` (`XLOPER12` by value). Ranges then arrive as values, as `ASFCells` expects. `U` (references) is only for functions that ask for them;
  - the type string's suffixes come from the decorator: `!` volatile, `$` thread-safe, `#` macro-equivalent. They can't be combined freely: `$` with `#` is refused, for example;
  - each argument description has at most 255 characters, as do the argument names together. The add-in cuts each description at 255 characters and the argument names after the last whole name that fits, with a line in the log.
- **Memory:** results come from the heap, are returned with `xlbitDLLFree` and are freed in `xlAutoFree12`, which Excel calls on the thread that got the result. Error results come from a table filled once at load and only read afterwards. So no result buffer is shared between threads, and thread-safe functions need no buffers of their own.
- **No Excel SDK files.** It declares `XLOPER12` itself and reaches `Excel12v` through `MdCallBack12`, the callback that Excel exports, as `xlcall.cpp` does. The structures are declared from Microsoft's documentation.

**The bridge to VBA**

- **One export, not three:** `AsfPy_Root(ByVal level As Long, ByRef out As Variant) As Long`, which returns a root proxy object (`py`). The rest goes through `IDispatch`: `py.call(...)`, `py.module(...)`, `py.lastError`. `level` is the access the add-in grants (3.4): the root of level 1 has no `module` member at all. So the limit is enforced in the `.xll`, not only in VBA. The single `Declare` keeps two VBA traps out:
  - `Declare` converts `String` arguments to ANSI, which loses non-Latin text. A `Variant` passed `ByRef` arrives as a `VARIANT*`, untouched;
  - a `Declare` can't return a `Variant` reliably, so the result comes back through an argument.
- **Finding the DLL.** `ASFPython.bas` declares `Lib "asfpy64.xll"` or `Lib "asfpy32.xll"` under `#If Win64`. VBA only calls `__stdcall` functions, so in the 32-bit build `AsfPy_Root` is `__stdcall` with a plain export name. A `Declare` finds the module only if Excel has already loaded it. Otherwise it fails with error 53. `ASFPython.bas` catches that error and shows "the Python add-in is not loaded", not the VBA error.
- **A proxy `IDispatch` for Python objects:**
  - `GetIDsOfNames` maps to `getattr`, and `Invoke` to a call or attribute read. `DISPID_VALUE` is the default value: the object's text (`x = obj`, `Debug.Print obj`). `DISPID_NEWENUM` gives `For Each` over iterables, through an `IEnumVARIANT` over `iter(obj)`. A Python iterator runs once, so `Reset` and `Clone` aren't supported;
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
  | number | `float` (`int` with a type hint; a fraction then gives `#VALUE!`) | number; `int` beyond 2^53 loses digits |
  | date | `float` (the cell holds a serial number); `date` or `datetime` with a type hint | number; a `date` or `datetime` as its serial number |
  | text, boolean | `str`, `bool` | value; text over 32,767 characters is cut |
  | empty | `None` | `""` |
  | error value | not called: the error is the result, as in `ASF_RUN` | — |
  | one row or column | `list` | one column |
  | several rows and columns | `list` of rows (numpy/pandas with a type hint) | grid, spilled |
  | `nan`, `inf` | — | `#NUM!` |
  | other object | kept in the handle cache | `<DataFrame#12>`, a handle other functions accept |
  | exception | — | `#VALUE!`; the reason via **Why this error?** |

  Phase 2 reads the type hints `int`, `float`, `str`, `bool`, `list`, `date` and `datetime` (also written as strings, and `list[...]` as `list`), for `@xl_func` functions and for `PY_CALL`. `str` makes a whole number `"12"`, not `"12.0"`; `list` makes one value a list of one, and an empty cell `[]`. An empty cell stays `None` under the other hints.

  From VBA (the `py` proxy), arrays arrive as `SAFEARRAY`s of `VARIANT`. ASF arrays have a lower bound of 1, ranges of 1, and VBA's `Array()` of 0. The proxy reads `LBound` and doesn't assume one. An ASF object (`ASF_Map`) is a VBA class with no type information, so `ASFPython.bas` converts it to a list of `[key, value]` pairs before the call, and Python gets a `dict`.
- **Handles are owned by their cell.** A handle is keyed by the calling cell (`xlfCaller`). When the cell is calculated again, its old object is dropped, so the cache doesn't grow with each recalculation. Handles are also scoped to the workbook.
- **Reload** re-imports the Python folder's modules and registers the functions again. Python itself stays (3.1), and so do the packages the modules imported (numpy is not imported again). Phase 2 makes it the add-in's command `AsfPyReload` (`Application.Run "AsfPyReload"`; a command, since Excel takes the C API calls of registration only from one). The modules of the folder are compiled from their source on each load, so a changed file is read again whatever its time stamp, and no `__pycache__` appears in the folder; a module that another module of the folder imported first runs once. Excel then recalculates by itself only the cells whose function changed slot; a cell whose function kept its slot keeps its old value until it is calculated again (Ctrl+Alt+F9), and so does a `PY_CALL` cell, since `PY_CALL` itself doesn't change (phase 2's checklist). The ribbon button **Reload Python** comes with phase 3, and it should then recalculate the workbook in full.
- **Errors.** Every failure keeps its exception: `py.lastError` (one line) and `py.lastTraceback`; `py.clearError()` forgets it. **Why this error?** clears it, calculates the cell again and shows what the cell's own call raised, as it does now for `ASF_RUN`. It must mark the cell dirty first (`Range.Dirty`): in Excel 16.0, `Range.Calculate` skips a cell that isn't dirty, and the cell's function is then not called at all (phase 2's checklist). `ASFRibbon.OnWhyError` calls `cell.Calculate` alone today. The button reads Python's error from phase 3, with the ribbon.
- **Settings live where ASF's do.** They go under the same registry key as `SaveSetting` (`HKCU\Software\VB and VBA Program Settings\ASF`): the Python folder, another Python than the add-in's own (`PythonVenv`, optional), and `PythonAccess` (3.4). The `.xll` reads that key with `RegGetValueW`. An administrator then has one place to set, or lock, both add-ins. There is no `asfpy.cfg`.
- **Pure Python where possible:** tested on Linux in CI, with a fake Excel in place of the `.xll`.

### 3.3 ASF integration

**Worksheet functions:**

```
=PY_CALL("stats.zscore", A2:A500)            stats.py in the Python folder, exported function zscore
=forecast(B2:B40, 12)                        a function registered with @xl_func
=ASF_CALL("summary", A2:A500)                summary.vas in the library, which asks for Python (below)
```

`PY_CALL(name, arg1, …, arg29)` takes up to 29 arguments; the ones left out at the end are dropped, one left out between two is `None`. An unknown name gives `#NAME?`, an exception `#VALUE!`, and an error value among the arguments is the result, as for `@xl_func` functions. `@xl_func` functions are exported too, as `module.function`.

**Library scripts** get a `py` object when they ask for it and the host allows it (3.4):

```js
// @asf python: call
return py.call('stats.zscore', $1);
```

```js
// @asf python: modules
let pd = py.module('pandas');
let df = pd.DataFrame($1);
return df.describe().values.tolist();
```

**Changes to the ASF engine,** small and useful beyond Python:

1. **Trusted injection.** A host method such as `InjectHostObject(name, obj)` injects an object that stays an object while `AppAccess` is off. Only VBA can call it, not scripts, so the sandbox still decides what a script can reach. The add-in uses it for `py`.
2. **The method name as written.** `TryCallHost` passes the name with its case to `CallByName`. COM objects ignore case, so nothing else changes.
3. Optionally, **keywords after a dot** (`obj.import`, `obj.default`), as JavaScript allows. The design uses `py.module` and doesn't depend on this change.

Phase 3 made 1 and 2: `ASF.InjectHostObject` (`ASF_Runtime.InjectHostObject`), and the name as written in `TryCallHost` and in the call of a host object's method from `DispatchMethod` (`METHOD_NAME_`), where it was lower-cased; reading a property already kept it. LibreOffice's `CallSpecial` matches UNO names without case, so it doesn't mind. Testing it in LibreOffice found two older gaps for Basic objects there, fixed with it: its `CallByName` takes `VbGet Or VbMethod` as neither (so `LOhelper.CallMember` tries a method call, then a property read), and a property read passed a missing optional argument, which LibreOffice raises on (it passes `Empty` now). `test/TestRunner.bas` checks both with `test/AsfFakeHost.cls`, a host object that runs in Excel and in LibreOffice (`libreoffice/tools/run_tests.py` loads it).

**In `ASF.xlam` (phase 3):** `ASFPython.bas` reads the `python` header and `PythonAccess` and gets `py` from the `.xll` (`AsfPy_Root`, declared only under `#If VBA7` on Windows, so the module also runs on Mac and in LibreOffice, where Python is refused with a message). `ASF_CALL` runs a script that asks for Python in an engine of its own, one for each level, holding `py`; the engine the worksheet functions share never holds it, since it also runs `ASF_RUN`. `ASF_RUN` with the header is refused. When a script with `py` fails, its message ends with Python's error, read from `py.lastError`, since `CallByName` drops it (2). **Run script** does the same for the scripts it runs. The ASF tab gets a **Python** group.

The proxy never hands a raw COM object to ASF. A Python function that returns Excel's `Application` gives a script a `#VALUE!`, not Excel access. The ribbon gets **Python folder** and **Reload Python**. **Why this error?** also shows the Python traceback. The descriptions go into `ASFRegister`, as they do now.

### 3.4 Security: trust follows where the code came from

ASF can sandbox a script. **Python cannot be sandboxed:** any Python code can read files, open sockets and start programs. So the design sorts code by where it comes from, not by what it calls:

| Comes from | Trusted? | Gets |
|---|---|---|
| The workbook: `ASF_RUN` script text and cell values | No: anyone can send a workbook, and formulas run without a macro warning | No `py`. Python from a formula goes through `PY_CALL` or `@xl_func` functions, which reach only installed functions |
| The library folder: `.vas` scripts the user installed | Yes, like any installed script | `py`, when the script asks for it and the host allows it (below) |
| The Python folder: `.py` modules the user installed | Yes | Everything Python can do |

**Two keys: the script asks, the host decides.** A script shows its intent with a header, and the host setting is the final word. Without both, `py` is undefined.

| `// @asf python:` in the script | `PythonAccess` setting of the host | Result |
|---|---|---|
| none | any | no `py` |
| `call` | `off` (the default) | the script doesn't run; the message names the setting it needs |
| `call` | `call` or `modules` | `py.call` of exported functions |
| `modules` | `off` or `call` | the script doesn't run; the message names the setting it needs |
| `modules` | `modules` | `py.call` and `py.module` |

- The setting is under ASF's registry key (3.2), so an administrator can set it or lock it.
- The add-in asks the `.xll` for the root of the granted level (3.1): level 1 for `call`, level 2 for `modules`. A level-1 script can't reach `module` by any path.
- A script that asks for more than the host allows is stopped with a clear message, not given less than it asked for. That way it never fails halfway through.
- `PY_CALL` and `@xl_func` functions don't go through ASF scripts. They work whenever `asfpy.xll` is loaded, since loading it is the host's decision, and they reach only installed, exported functions. `PythonAccess` doesn't change them.

**Why there is no package allow-list.** The first draft let `py.import` reach `numpy`, `pandas` and `scipy` from any formula. A list of packages doesn't hold:
- `numpy.load(f, allow_pickle=True)` and `pandas.read_pickle` run code from a file;
- `pandas.read_csv` reads any path or URL, and `DataFrame.to_csv` writes one;
- `numpy.ctypeslib.load_library` loads a DLL;
- modules reach other modules through their attributes.

Making that safe would mean reviewing every callable of every package, version by version. The table above avoids the problem instead: the workbook can only pick which installed function runs, and with what data.

**Rules:**

1. **No Python source in workbooks.** There is no `PY_EVAL("…")` and no Python in cells. Excel's own `=PY()` is different: it runs in Microsoft's cloud sandbox.
2. **`PY_CALL` looks names up; it doesn't import them.** `"stats.zscore"` is found in the registry built when the Python folder loaded. A cell string never reaches `importlib`, so `PY_CALL("os.system", …)` finds nothing.
3. **Exported functions receive untrusted data.** The documentation gives authors a checklist: no paths, URLs, `eval` or pickle taken from arguments, and a size limit on the data.
4. **The import path is pinned.** `sys.path` holds the Python folder, the `site-packages` of the Python (the add-in's own or the venv's) and the standard library. It never holds `''` or the current folder. Excel's current folder is often the folder of the workbook just opened, such as Downloads, and a `numpy.py` placed next to a workbook must not be imported. `safe_path` and ignoring the environment (`PYTHONPATH`, `PYTHONSTARTUP`) follow from the start-up option chosen in 3.1. DLLs are loaded by full path only (3.1).
5. **Proxies stay inside the rules.** Attributes that start with `_` are refused. Module objects are not handed out except through `py.module`, and raw COM objects never. The origin table is the real barrier; this is defence in depth.
6. **Handles can't cross workbooks** (3.2). A handle typed into a cell names nothing outside its own workbook.
7. **Distribution.** Excel has blocked `.xll` files from the internet since 2023 (Mark of the Web), and signing doesn't change that. Users install through a script that unblocks the file, or from a trusted location. Since the add-in carries its Python (3.1), installing is that one file. CI publishes the `.xll` with a checksum, and organisations can sign it for the "require signed add-ins" policy. Expect antivirus software to look hard at any unsigned `.xll` that loads Python: XLL files are a known malware vehicle.

### 3.5 Platforms

- **Excel for Windows, 64-bit and 32-bit:** the targets, Excel 2016 to Microsoft 365. Many Excel 2016 installs are 32-bit, so both builds ship (3.1). 32-bit support lasts as long as python.org publishes 32-bit Windows installers.
- **Excel for Mac:** an `.xll` cannot load there. Python would have to run in a separate process that VBA calls (as xlwings does), which is slower and a separate project. Mac keeps ASF without Python.
- **LibreOffice** (phase 5, built):
  - **One script in the extension.** `ASF.oxt` carries `python/asfpy_lo.py`, which `libreoffice/tools/build.py` makes from `libreoffice/python/asfpy_lo.py` and the text of `boot.py`: the same rules as in Excel (`@export`, levels, names without case, `_` refused). Basic reaches it through the script provider, at `vnd.sun.star.script:ASF.oxt|python|asfpy_lo.py$fn?language=Python&location=user:uno_packages` (the package's name comes from `PackageInformationProvider`), never with `location=document`: LibreOffice can embed Python in a document, and that comes with the document. The extension's manifest lists `python/` as `application/vnd.sun.star.framework-script`. LibreOffice's own Python runs it (3.10 on Windows), so `boot.start` takes `own_python=False` there: LibreOffice's output and `sys.path` stay as they are.
  - **`py` in ASF scripts.** `LOpython.PyRoot(level)` gives the root, and the host injects it (`engine.InjectHostObject "py", LOpython.PyRoot(1)`): the host decides, as `PythonAccess` does in Excel. Python objects reach Basic as a pyuno object that implements `XInvocation`, which Basic calls by name. The spike found three things:
    1. Basic's `TypeOf x Is AnyClass` is True for any object that implements `XInvocation` (LibreOffice's own Invocation adapter too; a plain UNO object gives False). The engine tells its maps by `TypeOf x Is ASF_Map`, so it never holds the proxy itself: `PyRoot` and `LOhelper.CallSpecial` wrap it in `ASF_PyObject`, a Basic class, and `CallSpecial` calls the proxy's `invoke`.
    2. pyuno doesn't take its own `XInvocation` object back as an argument (it asks for `XInvocation2`), so a Python object passed back to Python goes as a reference text that `asfpy_lo.py` turns back into the object.
    3. A failed call raises a `RuntimeException` with Python's message, which the engine reports; `py.lastError` has it too.
  - **`PY_CALL`** is a Basic function of `LOpython` in the `ASF` library (up to 15 arguments), for cells once the library is loaded. Calc passes a range as a 2-D array from 1, which goes to Python as rows. Two findings: arrays `ReDim`'d again into the same variable reach Python as `None` (each row is now built in a function of its own); and Calc takes no error value from a Basic function, so in a cell an unknown name gives the text `#NAME?` and an exception `#VALUE!`, while `LOpython.PyCall`, for Basic code, gives `CVErr(2029)` and `CVErr(2015)`. In a headless LibreOffice no Basic function runs from a cell, not even a one-line one (`#VALUE!`, whether the document is loaded with `MacroExecutionMode` 4 or not, from the API or from Basic): `libreoffice/tools/run_python.py` checks `PY_CALL` as Calc calls it, from Basic, and Calc cells are left to a check by hand.
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

Decision: write our own small add-in, and read xlOil and Excel-DNA for ideas, not code. The pywin32 server is not used, not even as a stand-in: phases 2 and 3 start once the `.xll` is done.

## 5. Phases

| Phase | Delivers | Done when |
|---|---|---|
| 0 | Investigation, design, toolchain spike | This page; the MinGW build |
| 1a | **Excel spike,** both bitnesses. The `.xll` loads a venv's Python (since then, its own by default) by full path, registers one slot function, and gives VBA a root proxy through `AsfPy_Root` | In Excel 2016 (32-bit and 64-bit) and Microsoft 365: a formula returns a value from Python; VBA calls `py.call`; the start-up option of 3.1 is chosen; the 32-bit slot stub works, or the fallback is chosen; the cost of one call and of start-up is measured; Excel closes cleanly with numpy loaded. **Built; passes in 64-bit Excel 16.0** |
| 1b | **The `.xll`, complete.** Start-up, both slot pools, registration, memory, threads and the GIL, the proxy, roots of both levels. Only a minimal Python bootstrap, enough to test it | The fake Excel host (below) runs both builds in CI; the struct-layout tests pass; a checklist in Excel 2016 (32-bit and 64-bit) and Microsoft 365 passes. **Built; the fake host passes natively in 64-bit and 32-bit; the checklist passes in 64-bit Excel 16.0; CI, 32-bit Excel and Excel 2016 are still to run** |
| 2 | Starts after phase 1. `PY_CALL`; `@xl_func` and `@export` with plain values and arrays; errors and **Why this error?**; reload; registry settings | Python tests on Linux; the `.xll` driven by a fake Excel host in CI (see below); a checklist run in Excel. **Built: `PY_CALL` in the `.xll`, type hints, `lastTraceback` and `clearError`, the command `AsfPyReload`; the settings were already ASF's (`ASF\AddIn`). `tests/unit/test_boot.py` and the fake host pass on Windows in 64-bit and 32-bit; the checklist passes in 64-bit Excel 16.0; CI, 32-bit Excel and Excel 2016 are still to run. The ribbon's buttons are phase 3's** |
| 3 | Starts after phase 1. Engine changes (3.3); `py` in library scripts; the two keys of 3.4 (`python` header and `PythonAccess`); ribbon buttons | Plus ASF test cases with a fake `py` object in LibreOffice and Excel. **Built: the engine's trusted injection and names as written, `ASFPython.bas`, the two keys in `ASF_CALL` and **Run script**, the ribbon's **Python folder**, **Reload Python** and **Why this error?** with Python. The engine's tests (`TestRunner`, with `AsfFakeHost`) and the add-in's checks (`run_lo.py`, `PythonAccess` off) pass in LibreOffice; the Excel checklist's phase 3 rows need `ASF.xlam` built again** |
| 4 | numpy and pandas conversions, object handles, `thread_safe`, async functions, `@xl_macro` | **Built.** Type hints `numpy.ndarray`, `pandas.DataFrame` (the first row the column names) and `pandas.Series`, found by module and name without importing either; arrays, DataFrames and Series back as cells. Handles: any other result, or a function hinted `-> object`, stays in Python as `<Type#n>`, owned by its cell (`xlfCaller` and `xlSheetNm`, read only when handles are in use) and scoped to its workbook. Asynchronous functions (`async def` on the add-in's event loop, `asynchronous=True` on its threads): registered as `>Q…X`, the result through `xlAsyncReturn` from the add-in's thread, with the GIL let go meanwhile. `@xl_macro`: commands (macro type 2) with arguments and a result, and `xl_app()`, Excel's `Application` through a workbook window's object (`AccessibleObjectFromWindow`), wrapped by pywin32 or comtypes, in commands only. `thread_safe` stays as phase 1b made it; handles work on calculation threads. `tests/unit/test_boot.py` (39 tests) and the fake host (103 checks 64-bit, 96 32-bit) pass, and so does the Excel checklist in 64-bit Excel 16.0 (with pandas and pywin32; its two failures were its own `Like` patterns) |
| 5 | LibreOffice `PY_CALL` and `py` (after its spike); later, if ever, the ribbon from Python, RTD and task panes | **Built** (3.5): `asfpy_lo.py` in `ASF.oxt`, `LOpython.bas` and `ASF_PyObject` in the `ASF` library. `libreoffice/tools/run_python.py` installs the extension in a temporary profile and passes its 28 checks; the engine's tests and the add-in's cases still pass. Calc cells need a check by hand; the ribbon from Python, RTD and task panes are left out |

Phases 2 and 3 already give users what matters most: Python functions in formulas and in ASF scripts.

**Testing the `.xll` without Excel.** `xlcall.cpp` finds `MdCallBack12` with `GetProcAddress(GetModuleHandle(NULL), …)`, that is, in the program that loaded the add-in. A small test program can export its own `MdCallBack12`, answer `xlfRegister` and the other calls the add-in makes, load `asfpy.xll` and call its slots. Built with MinGW, it runs under Wine with a Windows Python. **Confirmed in phase 1a, for both builds:** `tests/fakehost/fakehost.c`, with the results in `tests/fakehost/README.md`. Since phase 1b, `tests/fakehost/run_fakehost.py` also runs it natively on Windows, with the MSVC build, and `.github/workflows/python-addin.yml` runs both: MinGW under Wine on Linux, MSVC on Windows with 64-bit and 32-bit Pythons. Wine's C runtime lacks a function numpy 2.x uses (`crealf`), so the Wine job pins numpy 1.26 for its numpy checks. The GitHub-hosted Windows runners have no Excel, so the last check stays a self-hosted runner or a checklist run by hand before each release.

## 6. Open questions

Decided:
- 32-bit Excel is supported: two builds, `asfpy64.xll` and `asfpy32.xll` (3.1).
- Start-up: one of the two options of 3.1, chosen by the spike of phase 1a.
- Python in ASF scripts needs two keys: the script's `python` header and the host's `PythonAccess` setting (3.4).
- Phases 2 and 3 start after the `.xll` is done (phase 1). No pywin32 stand-in.
- `PY_CALL` and `@xl_func` functions work whenever the `.xll` is loaded; `PythonAccess` governs only `py` in ASF scripts.
- Installing: the `.xll` carries its Python and unpacks it on first load; pip runs from **Python packages**. `PythonVenv` is optional (3.1).
- The two keys are for Python only. Python's power can't be sandboxed or easily delegated to users; `appAccess` keeps working as it does now.

Open:
- **The name:** `asfpy64.xll` and `asfpy32.xll` beside `ASF.xlam` (proposed, as `ASF.xlam` must keep working without Python), or one `ASF.xll` one day.
