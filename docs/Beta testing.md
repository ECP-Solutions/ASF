---
layout: default
title: Beta testing
has_children: false
nav_order: 7
description: "How to test the Python add-in in Excel: a clean set-up, the checklist run, what to send back, and fixes for the usual problems."
---

# Beta testing the Python add-in
{: .no_toc }

Thank you for testing. This page takes you from a clean set-up to a result workbook to send back, in about 15 minutes. Follow the steps in order: most problems in earlier runs came from a file or a setting left over from a previous version.

1. TOC
{:toc}

---

## What you need

- Excel 2016 or later, or Microsoft 365, **for Windows**.
- Internet access the first time, to download the files and to install three Python packages (numpy, pandas, pywin32). Without it the checklist still runs, and skips the rows that need those packages.
- About 15 minutes, and Excel to yourself: the checklist opens workbooks, recalculates and unticks the add-in for a moment.

You don't need to install Python: the add-in carries its own.

---

## Step 1. Clean up an earlier test

Skip this step if you have never tested the Python add-in on this computer.

1. **Untick the old add-in.** In Excel: File > Options > Add-ins. At the bottom, choose **Manage: Excel Add-ins** and click **Go**. Untick every `asfpy64` or `asfpy32` entry and click **OK**.
2. **Clear the Python setting.** Earlier versions needed a venv, set with `AsfPySetup` (the setting `PythonVenv`). If it's still set, the add-in uses that Python instead of its own. Open the VBA editor (**Alt+F11**), then the Immediate window (**Ctrl+G**), paste this line and press Enter:
   ```vb
   On Error Resume Next: DeleteSetting "ASF", "AddIn", "PythonVenv"
   ```
   To check it's gone, paste this one: it should print `(not set)`.
   ```vb
   ?GetSetting("ASF", "AddIn", "PythonVenv", "(not set)")
   ```
3. **Close Excel** completely (all windows).
4. **Empty the Python folder.** In File Explorer's address bar, type `%APPDATA%\ASF\python` and press Enter. Delete everything in it, in particular:
   - `spike.py`: the old name of the test file. With it, and without the new `checks.py`, every check fails with `#NAME?`.
   - `zz_slots.py`: a file for automated tests only. It adds 1,000 functions and makes the checklist slow.
   - `aa_reload.py`, if an interrupted run left it.
5. **Delete old copies of the add-in** (`asfpy64.xll`, `asfpy32.xll`) wherever you saved them. Two copies in Excel at once give the alert "another copy of this add-in already runs Python".
6. **Optional, for a first-install test:** delete `%LOCALAPPDATA%\ASF\runtime` too. The add-in then unpacks its Python again on its first load, as on a new computer.

---

## Step 2. Download the files

Check Excel's bitness first: File > Account > **About Excel**. The first line ends with **64-bit** or **32-bit**.

Download with these links. They give the files themselves: don't use the browser's "Save page as", which saves a web page instead.

| File | For | Link |
|---|---|---|
| `asfpy64.xll` | 64-bit Excel | [download](https://github.com/ECP-Solutions/ASF/raw/main/excel/python/tests/excel/asfpy64.xll) |
| `asfpy32.xll` | 32-bit Excel | [download](https://github.com/ECP-Solutions/ASF/raw/main/excel/python/tests/excel/asfpy32.xll) |
| `checks.py` | the Python functions the checklist calls | [download](https://github.com/ECP-Solutions/ASF/raw/main/excel/python/tests/excel/checks.py) |
| `AsfPyChecks.bas` | the checklist | [download](https://github.com/ECP-Solutions/ASF/raw/main/excel/python/tests/excel/AsfPyChecks.bas) |
| `ASF.xlam` | optional: the rows for ASF scripts | [latest release](https://github.com/ECP-Solutions/ASF/releases/latest/download/ASF.xlam) |

Check the names in your Downloads folder. A browser that already had a file of that name saves `checks (1).py`; rename it to exactly `checks.py`. If Windows hides extensions, `checks.py.txt` shows as `checks.py`: turn on View > **File name extensions** in File Explorer to be sure.

---

## Step 3. Put the files in place

1. **The add-in:** make a folder of your own, such as `C:\ASF`, and move the `.xll` for your Excel there. Right-click it, choose **Properties**, tick **Unblock** if the box is there, and click **OK**. Excel refuses add-ins that Windows marks as downloaded.
2. **The Python folder:** move `checks.py` to `%APPDATA%\ASF\python` (create the `ASF` and `python` folders if they're missing). It must be the only `.py` file there.
3. **`ASF.xlam`** (optional): unblock it the same way and copy it to `%APPDATA%\Microsoft\AddIns`.

---

## Step 4. Load the add-in

1. Start Excel with a blank workbook.
2. File > Options > Add-ins > **Manage: Excel Add-ins** > **Go** > **Browse**. Pick the `.xll` in your folder and click **OK**. It appears ticked in the list; click **OK**.
   - With `ASF.xlam`: in the same dialog, tick **ASF** (or **Browse** to it).
3. The first load takes a few seconds: the add-in unpacks its Python into `%LOCALAPPDATA%\ASF\runtime`. No message means it started.

   If a message appears, read it: it says what's wrong (see [Problems and fixes](#problems-and-fixes)). The details are in `%TEMP%\asfpy.log`.

---

## Step 5. Check the packages

The add-in's Python already has numpy, pandas (64-bit only: pandas has no 32-bit build), openpyxl and pywin32, which the checklist uses. This step only checks that the package console opens.


1. Open the VBA editor (**Alt+F11**) and the Immediate window (**Ctrl+G**). Paste this line and press Enter:
   ```vb
   Application.Run "AsfPyPackages"
   ```
   With `ASF.xlam`, **Packages** in **Python studio** on the ASF tab does the same. Its **Python settings** button also shows and clears `PythonVenv` (step 1.2), and **About Python** shows which add-in and Python are loaded.
2. A black console window opens and lists the packages installed: numpy, pandas (64-bit), openpyxl and pywin32 among them. If one is missing, type `install` and its name after `pip>`, wait for `pip exited with 0`, and restart Excel afterwards.
3. Press Enter on the empty `pip>` line: the window closes.

In 32-bit Excel the pandas rows of the checklist say "skipped": that's expected.

---

## Step 6. Run the checklist

1. In Excel, keep multi-threaded calculation on: File > Options > Advanced > Formulas > **Enable multi-threaded calculation**.
2. Open the VBA editor (**Alt+F11**). Choose File > **Import File**, and pick `AsfPyChecks.bas`. It goes into the open workbook (`Book1`), not into `ASF.xlam`.
3. Put the cursor anywhere and run the macro **AsfPyTest**: in the Immediate window (**Ctrl+G**), type `AsfPyTest` and press Enter.
4. Wait without touching Excel. It takes from a few seconds to about a minute. A new workbook opens with two sheets:
   - **Checks:** one row per check, with `ok` or `FAIL` in column C, the timings, and in column D why a failed row failed;
   - **Formulas:** the formulas it calculated.
5. Read the first rows:
   - **Start-up of the Python folder** must list `checks.py` and `imported checks.py`. If the checklist stopped right there, `checks.py` isn't in the Python folder: go back to step 3.
   - **Python: the add-in's own (PythonVenv unset)** must say `yes`. If it says **PythonVenv's** instead, the setting from an earlier version is still there: go back to step 1.2, then restart Excel.
   - **Python: pip, for AsfPyPackages** must say `installed`.
6. The last rows give the number of failed checks and the time of the whole run.

---

## Step 7. Close Excel

Close Excel with the result workbook still open (numpy is loaded at this point). Excel should close without an error and without hanging. Start Excel again: the add-in loads by itself.

---

## Step 8. Send back

1. Save the result workbook (the one with the **Checks** sheet) as a normal `.xlsx`, such as `ASF_Python_<your Excel>.xlsx`.
2. Copy the log, `%TEMP%\asfpy.log`.
3. Send both, with:
   - your Excel's version and bitness (File > Account > About Excel; the first row of the sheet has it too) and Windows version;
   - anything unusual: a message, a hang, an Excel crash, a slow step.

If you can test more than one Excel (2016 32-bit, 2016 64-bit, Microsoft 365), send one workbook and one log for each.

---

## Problems and fixes

| What you see | What it means, and what to do |
|---|---|
| Excel says the file is not a valid add-in, or shows unreadable text | The `.xll` doesn't match Excel's bitness: take the other one (step 2) |
| Nothing happens when you tick the add-in, or Excel says the file is blocked | Windows still marks it as downloaded: **Unblock** it in Properties (step 3) |
| "Python started, but the Python folder gave no worksheet functions" | `checks.py` isn't in `%APPDATA%\ASF\python`, or it's misnamed (`checks (1).py`, `checks.py.txt`), or it's a saved web page: download it again with the link (step 2) |
| The checklist stops at once, saying `checks.py` isn't in the Python folder | The same: steps 2 and 3, then restart Excel |
| Most rows `FAIL` with `#NAME?` / `Error 2029`, and the run is very slow | The old `spike.py` is in the Python folder instead of `checks.py` (step 1.4). Each of thousands of calls then fails, which is what makes the run slow |
| "The Python folder gave N worksheet functions, and the add-in has 2000 slots" | Extra files in the Python folder, such as `zz_slots.py`: delete everything but `checks.py` |
| The `Python:` row says **PythonVenv's** | An earlier version's setting: clear it (step 1.2) and restart Excel |
| "another copy of this add-in already runs Python" | Two `.xll` files are loaded: untick all but one, delete the others, restart Excel |
| "another add-in has already loaded python3XX.dll" | Another Python add-in (PyXLL, xlOil) is loaded: untick it for the test |
| The package console doesn't open | Look for a message from Excel, and for the line `... didn't start` in `%TEMP%\asfpy.log` |
| `pip` fails to download | No internet, or a proxy: the rows for numpy, pandas and pywin32 are skipped; say so when you send back |
| Upgrading or removing a package fails with "access denied" | Excel has the package loaded: close Excel, start it again without running the checklist, and use the console then |
| A run-time error stops the checklist in the VBA editor | Note the error and the highlighted line, click **End**, and send them with the log |

**Start completely over:** untick the add-in, close Excel, and delete `%LOCALAPPDATA%\ASF\runtime`, `%APPDATA%\ASF\python` and `%TEMP%\asfpy.log`. Then follow the steps from step 1.2.

More about the add-in itself: [Python](Python.html).
