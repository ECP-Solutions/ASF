# Python inside Excel: the checks of phases 1a to 4

This folder is the Excel side of phases 1a to 4 of [the design](../DESIGN.md). `asfpy64.xll` and `asfpy32.xll` load a venv's Python into Excel and register the worksheet functions of `spike.py`. They also give VBA a `py` object. `AsfPySpike.bas` checks all of it and measures it. The phase 3 rows also need ASF.xlam loaded.

Phase 1a was the spike. Phase 1b completes the `.xll`: a pool of 1,000 slots, the volatile and thread-safe suffixes, functions unregistered when the add-in is unticked, `For Each` over Python objects, and an object's text as its default value.

Phase 2 adds `PY_CALL("module.function", arguments...)`, registered by the `.xll`; type hints (`int`, `float`, `str`, `bool`, `list`, `date`, `datetime`); `py.lastTraceback` and `py.clearError`, for **Why this error?**; and the command `AsfPyReload`, which imports the Python folder again without restarting Excel.

Phase 4 adds numpy and pandas type hints and results, object handles (`<Counter#1>` in a cell, the object for the functions that take the cell), asynchronous functions (`async def` or `asynchronous=True`, with `xlAsyncReturn`), and commands (`@xl_macro`, run with `Application.Run`, which may use `xl_app()`).

## What the checks must answer in Excel

| Question | Where the answer shows |
|---|---|
| Does Python start inside Excel 2016 (32-bit and 64-bit) and Microsoft 365? | `AsfPySpikeTest`: the first rows |
| Do the slots work, including the 32-bit stubs, for 0 to 7 arguments? | the `spike_*` rows, and Excel doesn't crash |
| Does `py` work from VBA through `CallByName`, as ASF will use it? | the `py.*` rows |
| Does `For Each` go through a Python object, and does `x = obj` give its text? | the `For Each` and `x = obj` rows |
| Do volatile and thread-safe functions work, and is `$` with `#` refused? | the `spike_volatile`, `spike_threadsafe` and `spike_refused` rows |
| Do the functions go when the add-in is unticked, and come back when it's ticked again? | the `Unticked` and `Ticked again` rows |
| Do `PY_CALL` and the type hints work in formulas, and does a cell's error come back with its traceback? | the `PY_CALL`, `spike_hints`, `spike_nextday` and `Why this error?` rows |
| Does a library script get `py` only with both keys, and does `ASF_RUN` never get it? (phase 3, with `ASF.xlam`) | the `ASF_CALL` and `ASF_RUN` rows at the end; `=ASF_CALL("py_check_modules")` gives `Which` only if the engine passes names as written |
| Do handles, asynchronous functions, numpy and pandas hints, and commands work? (phase 4) | the `spike_counter`, `spike_most`, `spike_async`, `spike_matrix`, `spike_frame`, `spike_describe`, `spike_macro` and `spike_app` rows. `spike_frame` needs pandas and `spike_app` pywin32 in the venv (`pip install pandas pywin32`); without them their rows say so |
| Does `AsfPyReload` read a module added, changed and removed? | the `AsfPyReload` rows (they write and delete `aa_reload.py` in the Python folder) |
| What does a call cost, and what does start-up cost? | the timing rows; the start-up lines of `%TEMP%\asfpy.log` |
| Does Excel close cleanly with numpy loaded? | step 9 |

## Steps

1. **Python of Excel's bitness.** Check Excel's bitness under File > Account > About Excel. Install a python.org Python 3.11 or later of the same bitness (the "Windows installer (64-bit)" or "(32-bit)"). Store Python and free-threaded builds don't work.
2. **A venv,** with numpy to try a real package:
   ```
   py -3.12 -m venv %USERPROFILE%\asfpy-venv
   %USERPROFILE%\asfpy-venv\Scripts\python -m pip install numpy
   ```
   For 32-bit, use `py -3.12-32`.
3. **The Python folder.** Copy `spike.py` to `%APPDATA%\ASF\python` (create the folder). Take the file itself: on GitHub, use **Download raw file**. A page saved with the browser's "Save as" is HTML, not Python. The name must be exactly `spike.py`: browsers sometimes save it as `spike (1).py` or `spike.py.txt`. When Excel loads the add-in, an alert says so if the folder gave no functions; the first rows of `AsfPySpikeTest` say why.
4. **The add-in.** Copy `asfpy64.xll` (64-bit Excel) or `asfpy32.xll` (32-bit Excel) to a folder of your own. In the file's Properties, tick **Unblock** if Windows shows it: Excel refuses add-ins downloaded from the internet.
5. **Settings.** Import `AsfPySpike.bas` into a new workbook (Alt+F11, File > Import File) and run `AsfPySetup`. Give it the venv folder and the Python folder.
6. **Load and test.** File > Options > Add-ins > Manage: Excel Add-ins > Go > Browse, then pick the `.xll`. Run `AsfPySpikeTest`. A new workbook shows each check, `ok` or `FAIL`, and the timings. Its second sheet holds the formulas. For the thread-safe rows, keep multi-threaded calculation on (File > Options > Advanced > Formulas).
7. **Phase 3: `ASF.xlam` built again,** with the engine and add-in of phase 3: `powershell -ExecutionPolicy Bypass -File excel/tools/build.ps1 -TemporaryTrustAccess` (see `excel/README.md`), then load it next to the `.xll`. `AsfPySpikeTest` then writes three scripts to the library folder, sets `PythonAccess` to `off`, `call` and `modules` in turn, and puts the setting and the folder back as they were. Without `ASF.xlam` it skips these rows.
8. **Phase 4** needs only the new `.xll` and `spike.py`, and `AsfPySpike.bas` imported again. For the pandas and `xl_app()` rows, `pip install pandas pywin32` in the venv first.
9. **Close Excel** with that workbook open (numpy loaded). Excel should close without an error or a hang. Then open Excel again: the add-in loads by itself.

Please send back the result sheet and `%TEMP%\asfpy.log` for each Excel you tried: 2016 32-bit, 2016 64-bit, Microsoft 365.

If Python doesn't start, Excel shows the reason when it loads the add-in, and `AsfPySpikeTest` repeats it. The log has the details.

**Changed a module of the Python folder?** Run `Application.Run "AsfPyReload"` (the Immediate window, Ctrl+G, takes it): the folder's modules are imported again and their functions registered again. Then press Ctrl+Alt+F9: Excel recalculates by itself only the cells whose function moved to another slot. **Changed the settings or the add-in? Close Excel and start it again.** Python starts once per Excel and stays until Excel closes, so unticking and ticking the add-in again registers the same functions again, without reading the Python folder. Browsing to a second copy of the `.xll` while one runs Python can't start another Python: the alert names the copy that runs.

## Results

| Excel | Add-in | Result |
|---|---|---|
| Microsoft Excel 16.0, 64-bit (Windows 10), Python 3.14.6 venv with numpy | asfpy64.xll of phase 1a | Passes: 11 worksheet functions and 14 exports registered; the only errors in the log are the ones the checks provoke |
| The same Excel, 4 calculation threads | asfpy64.xll of phase 1b (commit c2fc9c7) | 4 checks fail, the rest pass: `For Each`, `x = obj`, the suffixes, 20,000 thread-safe cells. Unticked, a function gave `#VALUE!`, not `#NAME?`: `xlfUnregister` leaves the name defined. The three refused calls raised error 440 "Automation error" through `CallByName`, without Python's text, while `py.lastError` held it: the checks now read `py.lastError` |
| The same Excel | phase 1b, the name deleted before `xlfUnregister` | 1 check fails: unticked, still `#VALUE!`. `CallByName` gets Python's message in `EXCEPINFO` (the log shows it) and still raises a bare 440; a late-bound call raises `-2147220503` (vbObjectError + 1001) with the message. So `CallByName` drops it, and ASF reads `py.lastError` |
| The same Excel | phase 1b, Excel-DNA's steps to remove a function | Excel took every step for all 14 functions (unregistered, registered again hidden, name deleted, hidden one unregistered), and an unticked function still gave `#VALUE!`: Excel keeps the name for the session. The add-in isn't called while unticked, so `xlAutoClose` is back to a plain `xlfUnregister`, and the check now asks for an error instead of 5.5 |
| The same Excel, 4 calculation threads | phase 1b as committed after this run | **All 34 checks pass.** `py.call` from VBA 8.5 µs; 2,000 cells of `spike_add` in 15 ms; unticked, `#VALUE!` |
| The same Excel | phase 2 | 45 checks pass: `PY_CALL` (with a range, and `#NAME?` for `os.system`), the type hints, a date back, three reloads (35–65 ms each, the log shows each module read again). 5 fail, all from `Range.Calculate`, which skipped cells that weren't dirty: **Why this error?** read no error, and after a reload the cells kept their old values. The checklist now marks them dirty first (`Range.Dirty`); its timings were 0 ms for the same reason |
| The same Excel | phase 2, cells marked dirty | **All 50 checks pass.** Why this error? reads the cell's `ValueError` and its traceback; after a reload a cell keeps its old value until it is calculated again, then shows the new code's; a removed export gives `#NAME?`. `py.call` from VBA 8.6 µs. The timings still read 0 ms: with automatic calculation `Range.Dirty` calculates at once, before the clock started, so the clock now starts before it |
| The same Excel, venv with numpy, pandas and pywin32 | phases 3 and 4, with `ASF.xlam` of phase 3 | 85 of 87 checks pass: the phase 3 rows; handles (`spike_most` reads the object, follows its cell's new one, and the old handle gives `LookupError`); 20 cells of `spike_async` in 0.8 s; `spike_matrix`, `spike_frame` and `spike_describe`; `Application.Run "spike_macro", 2, 3` gives 5, and `xl_app().ActiveSheet.Name` gives `Formulas`. The 2 that failed were the checklist's: VBA's `Like` takes `#` in `"<Counter#*>"` for a digit, so a handle never matched; the pattern is `"<Counter[#]*>"` now |

## Already checked without Excel

`tests/fakehost.c` stands in for Excel. It answers the C API calls the add-in makes, calls the registered functions through their slots and drives `py` through `IDispatch` as VBA does. Since phase 1b it also fills all 1,000 slots (with `tests/zz_slots.py`), checks the registered types, calls a thread-safe function from 8 threads at once, checks that 100,000 rounds of calls don't grow the process's memory, unticks and ticks the add-in again, and loads a second copy of it. Since phase 2 it calls `PY_CALL` with 30 arguments as Excel does, checks the type hints, and runs `AsfPyReload` with a module added, changed and removed. `tests/test_boot.py` tests `boot.py` alone, without Excel or Windows (22 tests). `tests/run_fakehost.py` runs it natively on Windows and under Wine on Linux.

Phase 2, on Windows 10 natively, with the MSVC build (`test_boot.py` passes with each Python too):

| Python | Add-in | Checks | One worksheet call | One `py.call` | Start-up, 1,016 functions |
|---|---|---|---|---|---|
| 3.14.6, 64-bit | asfpy64.xll | 82 of 82 | 5 µs | 5 µs | 650 ms |
| 3.14.6 venv with numpy 2.x, 64-bit | asfpy64.xll | 88 of 88 | 3 µs | 4 µs | 640 ms |
| 3.13, 64-bit (start-up through the global flags) | asfpy64.xll | 88 of 88 | 4 µs | 5 µs | 1,020 ms |
| 3.12.10, 32-bit (nuget.org's `pythonx86`) | asfpy32.xll | 82 of 82 | 6 µs | 6 µs | 1,440 ms |

Start-up compiles the folder's modules from their source since phase 2; most of it is `zz_slots.py`'s 1,000 functions.

Phase 4, the same way: `test_boot.py` has 39 tests (numpy and pandas ones skip without them). The fake host answers `xlfCaller`, `xlSheetNm` and `xlAsyncReturn` too, and checks handles (one cell's object for another, none across workbooks, the old one gone when its cell is calculated again), asynchronous functions (`async def` and a thread, an error argument, 30 calls at once), the type `>QX` of an asynchronous function, a command registered with macro type 2 and run, and `xl_app()` failing cleanly without Excel's window.

| Python | Add-in | Checks | One worksheet call |
|---|---|---|---|
| 3.14.6 venv with numpy 2.x, 64-bit | asfpy64.xll | 103 of 103 | 8 µs |
| 3.12.10, 32-bit (nuget.org's `pythonx86`) | asfpy32.xll | 96 of 96 | 12 µs |

Phase 1a, under Wine on a Linux container, with the MinGW build:

| Python | Add-in | Checks | One worksheet call | One `py.call` | Start-up |
|---|---|---|---|---|---|
| 3.11.9, 64-bit | asfpy64.xll | 30 of 30 | 2.0 µs | 2.3 µs | 390 ms |
| 3.12.10, 64-bit | asfpy64.xll | 30 of 30 | 2–3 µs | 2–4 µs | 250 ms |
| 3.12.10 venv with numpy 1.26.4, 64-bit | asfpy64.xll | 36 of 36 | 2.4 µs | 2.7 µs | 270 ms |
| 3.12.10, 32-bit | asfpy32.xll | 30 of 30 | 2.6 µs | 4.6 µs | 245 ms |
| 3.14.10, 64-bit | asfpy64.xll | 30 of 30 | 2.3 µs | 3.0 µs | 710 ms |
| 3.15.2, 64-bit | asfpy64.xll | 30 of 30 | 1.9 µs | 2.4 µs | 560 ms |

The times aren't Excel's: Excel adds its own cost per call, so they only show that the add-in's part is small. numpy 2.x doesn't load under Wine 9 (Wine's C runtime lacks `crealf`), which is a Wine limit; numpy 1.26 does. Natively, numpy 2.x works.

The GitHub workflow `.github/workflows/python-addin.yml` runs `test_boot.py` and these checks on every change to `excel/python`: the MinGW build under Wine with Pythons 3.11, 3.12 (with numpy 1.26), 3.14 and 32-bit 3.12, and the MSVC build natively with Pythons 3.11, 3.13 (with numpy), 3.14 and 32-bit 3.12.

## Building

On Windows with Visual Studio's C++ build tools, or anywhere with MinGW-w64:

```
python excel/python/xll/build.py                      # both add-ins; MinGW-w64 if it's on the PATH, else MSVC
python excel/python/tests/run_fakehost.py PYTHON...   # the checks; PYTHON is an install or a venv folder
python -m unittest discover -s excel/python/tests -p "test_*.py"   # boot.py alone, any OS
```

`build.py` writes to `excel/python/xll/build/`, and `run_fakehost.py` builds the fake host with the same toolchain. On Windows, `run_fakehost.py` sets the add-in's settings for the run and puts yours back afterwards. The add-ins in this folder are copies of the MSVC build.
