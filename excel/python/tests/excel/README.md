# The Python add-in in Excel: the checklist

Excel itself can't run in CI, so this folder holds the checks that run in Excel, by hand:

| File | What it is |
|---|---|
| `AsfPyChecks.bas` | The checklist: a VBA module whose `AsfPyTest` checks the add-in and measures it, and writes the results to a new workbook |
| `checks.py` | The Python side: the worksheet functions, commands and exported functions the checklist calls. The fake host (`../fakehost`) uses it too |
| `asfpy64.xll`, `asfpy32.xll` | The add-ins to check, built from this branch with the Python they carry |

## What the checklist answers

| Question | Where the answer shows |
|---|---|
| Does Python start inside Excel 2016 (32-bit and 64-bit) and Microsoft 365? | the first rows |
| Is it the add-in's own Python, unpacked under `%LOCALAPPDATA%\ASF\runtime`, with pip for its packages? | the `Python:` rows (with `PythonVenv` set, the row shows that Python instead) |
| Do the slots work, including the 32-bit stubs, for 0 to 7 arguments? | the `check_*` rows, and Excel doesn't crash |
| Does `py` work from VBA through `CallByName`, as ASF uses it? | the `py.*` rows |
| Does `For Each` go through a Python object, and does `x = obj` give its text? | the `For Each` and `x = obj` rows |
| Do volatile and thread-safe functions work, and is `$` with `#` refused? | the `check_volatile`, `check_threadsafe` and `check_refused` rows |
| Do the functions go when the add-in is unticked, and come back when it's ticked again? | the `Unticked` and `Ticked again` rows |
| Do `PY_CALL` and the type hints work in formulas, and does a cell's error come back with its traceback? | the `PY_CALL`, `check_hints`, `check_nextday` and `Why this error?` rows |
| Does a library script get `py` only with both keys, and does `ASF_RUN` never get it? (with `ASF.xlam`) | the `ASF_CALL` and `ASF_RUN` rows at the end; `=ASF_CALL("py_check_modules")` gives `Which` only if the engine passes names as written |
| Do handles, asynchronous functions, numpy and pandas hints, and commands work? | the `check_counter`, `check_most`, `check_async`, `check_matrix`, `check_frame`, `check_describe`, `check_macro` and `check_app` rows. `check_matrix` needs numpy, `check_frame` pandas and `check_app` pywin32 (step 2); without them their rows say so |
| Does `AsfPyReload` read a module added, changed and removed? | the `AsfPyReload` rows (they write and delete `aa_reload.py` in the Python folder) |
| What does a call cost, and what does start-up cost? | the timing rows; the start-up lines of `%TEMP%\asfpy.log` |
| Does a tooltip show a Python function's arguments as a formula is typed? | step 7 |
| Do macros run from Python studio with their arguments, and can they show Excel's InputBox and a message box? | step 8 |
| Does Excel close cleanly with numpy loaded? | step 9 |

## Steps

The full guide for testers, with downloads, clean-up of earlier runs and fixes for the usual problems, is [Beta testing](../../../../docs/Beta%20testing.md) in the documentation. In short:

1. **Clean up an earlier test.** Untick old `asfpy*.xll` entries, clear `PythonVenv` (Immediate window: `On Error Resume Next: DeleteSetting "ASF", "AddIn", "PythonVenv"`), close Excel, empty `%APPDATA%\ASF\python` (no `spike.py`, no `zz_slots.py`), delete old copies of the `.xll`.
2. **Download** `asfpy64.xll` or `asfpy32.xll` (Excel's bitness: File > Account > About Excel), `checks.py` and `AsfPyChecks.bas` as raw files, with exactly those names.
3. **Put them in place.** Unblock the `.xll` (Properties) and keep it in a folder of your own; `checks.py` alone in `%APPDATA%\ASF\python`. `ASF.xlam` is optional: with it, the checklist also checks ASF scripts with Python.
4. **Load the add-in:** File > Options > Add-ins > Manage: Excel Add-ins > Go > Browse. The first load unpacks the add-in's Python into `%LOCALAPPDATA%\ASF\runtime`.
5. **Packages:** none to install: the add-in's Python has numpy, pandas (64-bit only), openpyxl and pywin32. `Application.Run "AsfPyPackages"` lists them.
6. **Run:** import `AsfPyChecks.bas` into a blank workbook and run `AsfPyTest`, multi-threaded calculation on. The first rows must show `imported checks.py`, `Python: the add-in's own ... yes` and `pip ... installed`.
7. **Help as you type:** in a cell, type `=check_add(` then `1, `: a tooltip under the caret shows `check_add(a, b)` with `a`, then `b`, in bold, and `Adds two numbers.`; it goes on Esc or Enter. Typing `=check_ad` shows it beside Excel's list; in the formula bar too. Insert Function (Shift+F3) on `check_add` shows each argument's help.
8. **Macros from Python studio** (`Application.Run "AsfPyStudio"`, or **Python studio** on the ASF tab): in **Functions**, ▶ `check_macro` asks for `a` and `b`: 2 and 3 give a box saying 5; leaving `b` blank gives `check_macro failed` with Python's `TypeError`. ▶ `check_dialogs` with **Bring Excel to the front** ticked: Excel comes forward with its InputBox; 7 gives a message box over Excel saying `7 squared is 49`, and the studio then shows 49 (Cancel: `cancelled`).
9. **Close Excel** with the result open (numpy loaded): no error, no hang. Then open Excel again: the add-in loads by itself.

To check another Python instead (a venv, or a python.org install of Excel's bitness, 3.11 or later), run `AsfPySetup` and give its folder; leave the folder empty to go back to the add-in's own Python. Packages then go into that Python, with its own pip.

Please send back the result sheet and `%TEMP%\asfpy.log` for each Excel you tried: 2016 32-bit, 2016 64-bit, Microsoft 365.

If Python doesn't start, Excel shows the reason when it loads the add-in, and the checklist repeats it. The log has the details.

**Changed a module of the Python folder?** Run `Application.Run "AsfPyReload"`: the folder's modules are imported again and their functions registered again. Then press Ctrl+Alt+F9: Excel recalculates by itself only the cells whose function moved to another slot. **Changed the settings or the add-in? Close Excel and start it again.** Python starts once per Excel and stays until Excel closes, so unticking and ticking the add-in again registers the same functions again, without reading the Python folder. Browsing to a second copy of the `.xll` while one runs Python can't start another Python: the alert names the copy that runs.

## Results

Runs in Excel, newest last. Until the add-in carried its own Python, the checklist ran with a venv; the functions of `checks.py` were then in `spike.py`, as `spike_*`.

| Excel | Add-in | Result |
|---|---|---|
| Microsoft Excel 16.0, 64-bit (Windows 10), Python 3.14.6 venv with numpy | commit 8699144 | Passes: 11 worksheet functions and 14 exports registered; the only errors in the log are the ones the checks provoke |
| The same Excel, 4 calculation threads | commit c2fc9c7 | 4 checks fail, the rest pass: `For Each`, `x = obj`, the suffixes, 20,000 thread-safe cells. Unticked, a function gave `#VALUE!`, not `#NAME?`: `xlfUnregister` leaves the name defined. The three refused calls raised error 440 "Automation error" through `CallByName`, without Python's text, while `py.lastError` held it: the checks now read `py.lastError` |
| The same Excel | c2fc9c7, the name deleted before `xlfUnregister` | 1 check fails: unticked, still `#VALUE!`. `CallByName` gets Python's message in `EXCEPINFO` (the log shows it) and still raises a bare 440; a late-bound call raises `-2147220503` (vbObjectError + 1001) with the message. So `CallByName` drops it, and ASF reads `py.lastError` |
| The same Excel | c2fc9c7, Excel-DNA's steps to remove a function | Excel took every step for all 14 functions (unregistered, registered again hidden, name deleted, hidden one unregistered), and an unticked function still gave `#VALUE!`: Excel keeps the name for the session. The add-in isn't called while unticked, so `xlAutoClose` is back to a plain `xlfUnregister`, and the check now asks for an error instead of 5.5 |
| The same Excel, 4 calculation threads | commit d7c53bf | **All 34 checks pass.** `py.call` from VBA 8.5 µs; 2,000 cells of `check_add` in 15 ms; unticked, `#VALUE!` |
| The same Excel | commit ea28f45, before its last fix | 45 checks pass: `PY_CALL` (with a range, and `#NAME?` for `os.system`), the type hints, a date back, three reloads (35–65 ms each, the log shows each module read again). 5 fail, all from `Range.Calculate`, which skipped cells that weren't dirty: **Why this error?** read no error, and after a reload the cells kept their old values. The checklist now marks them dirty first (`Range.Dirty`); its timings were 0 ms for the same reason |
| The same Excel | commit ea28f45 | **All 50 checks pass.** Why this error? reads the cell's `ValueError` and its traceback; after a reload a cell keeps its old value until it is calculated again, then shows the new code's; a removed export gives `#NAME?`. `py.call` from VBA 8.6 µs. The timings still read 0 ms: with automatic calculation `Range.Dirty` calculates at once, before the clock started, so the clock now starts before it |
| The same Excel, venv with numpy, pandas and pywin32 | commit 7466591, with `ASF.xlam` 4.1.0 | 85 of 87 checks pass: the library script rows; handles (`check_most` reads the object, follows its cell's new one, and the old handle gives `LookupError`); 20 cells of `check_async` in 0.8 s; `check_matrix`, `check_frame` and `check_describe`; `Application.Run "check_macro", 2, 3` gives 5, and `xl_app().ActiveSheet.Name` gives `Formulas`. The 2 that failed were the checklist's: VBA's `Like` takes `#` in `"<Counter#*>"` for a digit, so a handle never matched; the pattern is `"<Counter[#]*>"` now |

## The add-ins in this folder

`asfpy64.xll` and `asfpy32.xll` here are MinGW-w64 builds (`python excel/python/xll/build.py mingw`), carrying CPython 3.14.10 and its packages (`xll/runtime.json`), the same files as the latest build kept in `excel/python/xll/build/`. The fake host passes with both (`../fakehost/README.md`). The builds that passed the checklist above in Excel were MSVC builds; `python excel/python/xll/build.py msvc` on Windows makes those, with the same Python.
