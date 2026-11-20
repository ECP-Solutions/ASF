# Phase 1a spike: Python inside Excel

This folder is the Excel side of phase 1a of [the design](../DESIGN.md). `asfpy64.xll` and `asfpy32.xll` load a venv's Python into Excel and register the worksheet functions of `spike.py`. They also give VBA a `py` object. `AsfPySpike.bas` checks all of it and measures it. It's a spike: one log file, no ribbon, no ASF integration yet.

## What the spike must answer in Excel

| Question | Where the answer shows |
|---|---|
| Does Python start inside Excel 2016 (32-bit and 64-bit) and Microsoft 365? | `AsfPySpikeTest`: the first rows |
| Do the slots work, including the 32-bit assembly stubs, for 0 to 7 arguments? | the `spike_*` rows, and Excel doesn't crash |
| Does `py` work from VBA through `CallByName`, as ASF will use it? | the `py.*` rows |
| What does a call cost, and what does start-up cost? | the timing rows; the start-up lines of `%TEMP%\asfpy.log` |
| Does Excel close cleanly with numpy loaded? | step 7 |

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
6. **Load and test.** File > Options > Add-ins > Manage: Excel Add-ins > Go > Browse, then pick the `.xll`. Run `AsfPySpikeTest`. A new workbook shows each check, `ok` or `FAIL`, and the timings. Its second sheet holds the formulas.
7. **Close Excel** with that workbook open (numpy loaded). Excel should close without an error or a hang. Then open Excel again: the add-in loads by itself.

Please send back the result sheet and `%TEMP%\asfpy.log` for each Excel you tried: 2016 32-bit, 2016 64-bit, Microsoft 365.

If Python doesn't start, Excel shows the reason when it loads the add-in, and `AsfPySpikeTest` repeats it. The log has the details.

## Already checked without Excel

`tests/fakehost.c` stands in for Excel. It answers the C API calls the add-in makes, calls the registered functions through their slots, and drives `py` through `IDispatch` as VBA does. `tests/run_wine.py` runs it under Wine on Linux. Results:

| Python | Add-in | Checks | One worksheet call | One `py.call` | Start-up |
|---|---|---|---|---|---|
| 3.11.9, 64-bit | asfpy64.xll | 30 of 30 | 2.0 µs | 2.3 µs | 390 ms |
| 3.12.10, 64-bit | asfpy64.xll | 30 of 30 | 2–3 µs | 2–4 µs | 250 ms |
| 3.12.10 venv with numpy 1.26.4, 64-bit | asfpy64.xll | 36 of 36 | 2.4 µs | 2.7 µs | 270 ms |
| 3.12.10, 32-bit | asfpy32.xll | 30 of 30 | 2.6 µs | 4.6 µs | 245 ms |
| 3.14.10, 64-bit | asfpy64.xll | 30 of 30 | 2.3 µs | 3.0 µs | 710 ms |
| 3.15.2, 64-bit | asfpy64.xll | 30 of 30 | 1.9 µs | 2.4 µs | 560 ms |

The times are Wine's on a Linux container, not Excel's. Excel adds its own cost per call, so they only show that the add-in's part is small. numpy 2.x doesn't load under Wine 9 (Wine's C runtime lacks `crealf`), which is a Wine limit; numpy 1.26 does.

## Building

```
python3 excel/python/xll/build.py                      # both add-ins, with MinGW-w64
python3 excel/python/tests/run_wine.py PYTHON...       # the checks under Wine
```

`build.py` writes to `excel/python/xll/build/`. The add-ins in this folder are copies of that output.
