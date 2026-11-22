# The fake Excel host

`fakehost.c` stands in for Excel. It answers the C API calls the add-in makes, calls the registered functions through their slots and drives `py` through `IDispatch` as VBA does. It also fills all 2,000 slots (with `zz_slots.py` and the copy `run_fakehost.py` makes of it, 2,000 functions), checks the registered types, calls a thread-safe function from 8 threads at once, checks that 100,000 rounds of calls don't grow the process's memory, unticks and ticks the add-in again, and loads a second copy of it. It calls `PY_CALL` with 30 arguments as Excel does, checks the type hints, and runs `AsfPyReload` with a module added, changed and removed. It also answers `xlfCaller`, `xlSheetNm` and `xlAsyncReturn`, and checks handles (one cell's object for another, none across workbooks, the old one gone when its cell is calculated again), asynchronous functions (`async def` and a thread, an error argument, 30 calls at once), the type `>QX` of an asynchronous function, commands registered with macro type 2 and run (the add-in's `AsfPyReload` and `AsfPyPackages`, and `@xl_macro` functions), and `xl_app()` failing cleanly without Excel's window. `run_fakehost.py` runs it natively on Windows and under Wine on Linux.

Natively on Windows 10, with the MSVC build of commit ea28f45 (fewer checks then):

| Python | Add-in | Checks | One worksheet call | One `py.call` | Start-up, 1,016 functions |
|---|---|---|---|---|---|
| 3.14.6, 64-bit | asfpy64.xll | 82 of 82 | 5 µs | 5 µs | 650 ms |
| 3.14.6 venv with numpy 2.x, 64-bit | asfpy64.xll | 88 of 88 | 3 µs | 4 µs | 640 ms |
| 3.13, 64-bit (start-up through the global flags) | asfpy64.xll | 88 of 88 | 4 µs | 5 µs | 1,020 ms |
| 3.12.10, 32-bit (nuget.org's `pythonx86`) | asfpy32.xll | 82 of 82 | 6 µs | 6 µs | 1,440 ms |

Start-up compiles the folder's modules from their source; most of it is `zz_slots.py`'s 1,000 functions. `zz_slots.py` is for the fake host only: in Excel's Python folder it only makes the checklist slower.

The same way, with the MSVC build of commit 7466591:

| Python | Add-in | Checks | One worksheet call |
|---|---|---|---|
| 3.14.6 venv with numpy 2.x, 64-bit | asfpy64.xll | 103 of 103 | 8 µs |
| 3.12.10, 32-bit (nuget.org's `pythonx86`) | asfpy32.xll | 96 of 96 | 12 µs |

Under Wine on a Linux container, with the MinGW build of commit 8699144 (30 checks then):

| Python | Add-in | Checks | One worksheet call | One `py.call` | Start-up |
|---|---|---|---|---|---|
| 3.11.9, 64-bit | asfpy64.xll | 30 of 30 | 2.0 µs | 2.3 µs | 390 ms |
| 3.12.10, 64-bit | asfpy64.xll | 30 of 30 | 2–3 µs | 2–4 µs | 250 ms |
| 3.12.10 venv with numpy 1.26.4, 64-bit | asfpy64.xll | 36 of 36 | 2.4 µs | 2.7 µs | 270 ms |
| 3.12.10, 32-bit | asfpy32.xll | 30 of 30 | 2.6 µs | 4.6 µs | 245 ms |
| 3.14.10, 64-bit | asfpy64.xll | 30 of 30 | 2.3 µs | 3.0 µs | 710 ms |
| 3.15.2, 64-bit | asfpy64.xll | 30 of 30 | 1.9 µs | 2.4 µs | 560 ms |

The times aren't Excel's: Excel adds its own cost per call, so they only show that the add-in's part is small. numpy 2.x doesn't load under Wine 9 (Wine's C runtime lacks `crealf`), which is a Wine limit; numpy 1.26 does. Natively, numpy 2.x works.

The add-in's own Python and the CI's Pythons, under Wine 9 on a Linux container with the MinGW build, 2,000 slots, in a fresh Wine prefix (the first load unpacked the add-in's Python in about 450 ms; pip in that Python installed a package from PyPI into its `Lib\site-packages`):

| Python | Add-in | Checks |
|---|---|---|
| the add-in's own, 3.14.10, 64-bit | asfpy64.xll | 100 of 100 |
| the add-in's own, 3.14.10, 32-bit | asfpy32.xll | 100 of 100 |
| 3.11.9, 64-bit | asfpy64.xll | 100 of 100 |
| 3.12.10 venv with numpy 1.26.4, 64-bit | asfpy64.xll | 107 of 107 |
| 3.14.10, 64-bit | asfpy64.xll | 100 of 100 |
| 3.12.10, 32-bit | asfpy32.xll | 100 of 100 |

Two things the fake host must do as Excel does, or Wine fails where Excel doesn't:
- **Stack above a call.** A 64-bit slot reads all 32 argument places from its caller's stack, whatever the function takes. Excel always calls from deep in its own frames; the fake host's threads keep 1 KB of room above their calls for that. Without it, a fresh thread under Wine read past the top of its stack.
- **Wine's runtime folder.** Wine 9 crashes on its first command without `/run/user/<uid>`, which a CI runner, having no login session, lacks: the workflow makes it, and `run_fakehost.py` says so when it's missing.

The GitHub workflow `.github/workflows/python-addin.yml` runs `../unit/test_boot.py` and these checks on every change to `excel/python`: the MinGW build under Wine with Pythons 3.11, 3.12 (with numpy 1.26), 3.14 and 32-bit 3.12, and the MSVC build natively with Pythons 3.11, 3.13 (with numpy), 3.14 and 32-bit 3.12.

## Running it

```
python excel/python/xll/build.py                              # both add-ins, first
python excel/python/tests/fakehost/run_fakehost.py            # with the Python each add-in carries
python excel/python/tests/fakehost/run_fakehost.py PYTHON...  # with installs or venvs instead
```

Without arguments, `run_fakehost.py` runs `bundled64` and `bundled32`: each add-in with the Python it carries, `PythonVenv` unset. A `PYTHON` folder is an install, a venv, or the tools folder of nuget.org's `python` or `pythonx86` package, run with the add-in of its bitness. It builds the fake host with the toolchain `build.py` used. On Windows it sets the add-in's settings for the run and puts yours back afterwards, and the log and the unpacked Python go to a temporary folder.
