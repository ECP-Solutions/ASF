---
layout: default
title: Add-ins
has_children: false
nav_order: 5
description: "ASF as an Excel add-in and a LibreOffice extension: installation, worksheet functions, the script library, the ASF tab and settings."
---

# ASF add-ins for Excel and LibreOffice
{: .no_toc }

ASF comes ready to use as an add-in, so workbooks can run ASF scripts without importing any VBA:

- **`ASF.xlam`**, the Excel add-in: worksheet functions (`ASF_RUN`, `ASF_CALL`, regex functions), a script library, an ASF tab on the ribbon, and `NewEngine()` for macros. Excel 2016 and later, on Windows and Mac.
- **`ASF.oxt`**, the LibreOffice extension: the same engine as a Basic library named `ASF`. LibreOffice 7.5 and later.

Python is a separate, optional add-in for Excel for Windows: see [Python](Python.html).

1. TOC
{:toc}

---

## Installation

Download the files from the [latest release](https://github.com/ECP-Solutions/ASF/releases/latest).

### Excel for Windows

1. Download [`ASF.xlam`](https://github.com/ECP-Solutions/ASF/releases/latest/download/ASF.xlam).
2. **Unblock it.** Excel blocks the macros of files downloaded from the internet. Right-click the file, choose **Properties**, tick **Unblock** and click **OK**. Windows doesn't show the box when there is nothing to unblock.
3. **Put it in a trusted location.** Copy it to `%APPDATA%\Microsoft\AddIns`: type that path in File Explorer's address bar. Excel trusts that folder by default. Any folder added under File > Options > Trust Center > Trust Center Settings > Trusted Locations works too.
4. **Load it.** In Excel: File > Options > Add-ins. At the bottom, choose **Manage: Excel Add-ins** and click **Go**. Tick **ASF**. If it isn't in the list, click **Browse** and pick the file.
5. **Check it.** The ribbon has an **ASF** tab, and in any cell `=ASF_RUN("return 6 * 7;")` gives `42`.

Excel loads the add-in every time it starts until you untick it. To update it, close Excel, replace the file and start Excel again.

### Excel for Mac

1. Download `ASF.xlam` and put it in a folder of your own.
2. In Excel: Tools > Excel Add-ins > **Browse**, and pick `ASF.xlam`.
3. Enable its macros when Excel asks.

### Signed VBA only

The add-in is not signed. An organisation that allows only signed VBA can sign the project with its own certificate (VBA editor: Tools > Digital Signature) before deploying it.

### Where it doesn't work

Excel for the web, iPad and Teams don't run VBA. Workbooks that use the functions show `#NAME?` there.

### LibreOffice

1. Download [`ASF.oxt`](https://github.com/ECP-Solutions/ASF/releases/latest/download/ASF.oxt).
2. In LibreOffice: Tools > Extension Manager > **Add...**, pick the file, and restart LibreOffice. From a terminal, `unopkg add ASF.oxt` does the same.
3. For `@(...)` VBA expressions only, install the [VBA Expressions extension](https://extensions.libreoffice.org/en/extensions/show/70059) 1.2.1 or later the same way.

See [LibreOffice](#libreoffice-1) below for using it.

---

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

The regular expressions are ASF's own engine, with JavaScript syntax, so results are the same on Windows and Mac.

The functions are in the **ASF** category of Insert Function (fx), with a description of each argument. **Ctrl+A** after `=ASF_RUN` opens the Function Arguments dialog, and **Ctrl+Shift+A** types the argument names in. The free [Excel-DNA IntelliSense](https://github.com/Excel-DNA/IntelliSense) add-in also shows their tips while you type: load its `.xll` next to `ASF.xlam`.

### Arguments

- A single cell gives its value. Numbers become doubles, as numbers in scripts are.
- One column or one row gives a flat array from 1: with `B2:B6`, `$1[1]` is B2.
- A range of several rows and columns gives an array of rows, each row an array from 1: `$1[row][column]`.
- An empty cell is `undefined`; write `x ?? ''` when you want text.
- An error value as a whole argument (`#N/A` in a single cell) is the result, without running the script, as Excel's own functions do. Errors inside a range arrive as their text (`"#N/A"`).
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

In Microsoft 365 and Excel 2021 an array result spills from the formula's cell. In Excel 2016 and 2019, select the output range, type the formula and press Ctrl+Shift+Enter: cells beyond the result get `""`, and a larger result is cut. The same workbook works in both.

**Why this error?** on the ASF tab tells why a cell gave `#VALUE!`.

---

## The script library

`ASF_CALL("name", ...)` runs `name.vas` from the library folder. That folder is also the only place `import` may read from, for `ASF_CALL` and `ASF_RUN` alike. **Library folder** on the ASF tab opens it.

- Windows: `%APPDATA%\ASF\scripts`
- Mac: `~/Library/Group Containers/UBF8T346G9.Office/ASF/scripts`, where Excel for Mac can read without asking

A library script is read again when its file changes. Scripts are read as ANSI text, so save them as ASCII or Windows-1252.

A script may start with a header of comment lines. It tells **Run script** how to list and run it:

```
// @asf name: Apply discount rules
// @asf appAccess: true
// @asf args: selection
// @asf python: call
```

| Key | Meaning |
|---|---|
| `name` | The name **Run script** shows; the file name when there is none |
| `appAccess` | `true`: the script may use Office objects. Default `false` |
| `args` | `selection`: `$1` is the selected range. Default: no arguments |
| `python` | `call` or `modules`: the script gets `py` from the Python add-in. `ASF_CALL` reads this key too. See [Python](Python.html) |

With `appAccess: true` and `args: selection`, `$1` is the `Range` itself, so the script can write cells, format them or reach the workbook. Only library scripts run from **Run script** can have Office access.

---

## The ASF tab

| Button | What it does |
|---|---|
| **Run script** | Lists the scripts of the library folder (the first 30) and runs the one you pick, then shows its result or why it stopped |
| **Library folder** | Opens the library folder, made first when it is missing |
| **Why this error?** | Calculates the active cell again and shows why its `ASF_RUN` or `ASF_CALL` gave `#VALUE!`. With the Python add-in, also the Python error of a `PY_CALL` or Python function |
| **Reset engine** | Drops the engine of the worksheet functions and their compiled scripts; the next formula reads the settings again |
| **Python folder**, **Reload Python**, **Python packages** | For the Python add-in: see [Python](Python.html) |
| **Help** | Opens the add-in's README on GitHub |

---

## Settings

| Setting | Default |
|---|---|
| `LibraryFolder` | the folder above |
| `RunEnabled` | `1`. With `0`, `ASF_RUN` gives `#VALUE!` and only library scripts run |
| `StepLimit` | `10000000` loop iterations and calls per script; `0` for no limit |
| `PythonAccess` | `off`. What library scripts that ask for Python may do: `call` or `modules`. See [Python](Python.html) |
| `PythonVenv`, `PythonFolder` | Read by the Python add-in. See [Python](Python.html) |

The settings are kept with `SaveSetting`: on Windows in `HKEY_CURRENT_USER\Software\VB and VBA Program Settings\ASF\AddIn`, where an administrator can set them by group policy; on Mac in a preferences file. Change one from the Immediate window of the VBA editor (Alt+F11, then Ctrl+G):

```vb
Application.Run "ASF.xlam!ASFSettings.ChangeSetting", "RunEnabled", "0"
```

---

## Security

A workbook's formulas run when it recalculates, with no macro warning: the add-in is trusted, and the scripts are only data. So every script the functions run is limited, wherever it comes from:

- **No Office objects.** `AppAccess` is off and ranges arrive as values.
- **Imports only from the library folder.** An import with `../` or an absolute path outside it gives `#VALUE!`.
- **A step limit.** `while (true) { }` in a cell can't freeze Excel.
- **`ASF_CALL` only runs installed scripts.** It only takes bare script names, so a received workbook can only run scripts you installed.
- **Scripts are never stored inside workbooks,** and `ASF_RUN` can be turned off.
- **Office access only for library scripts run from the ribbon,** and only when their header asks for it.
- **Python only for library scripts, with two keys:** the script's `python` header and the `PythonAccess` setting.

---

## Macros in other workbooks

Add a reference to the add-in (VBA editor: Tools > References > **ASFAddIn**), then:

```vb
Dim e As ASF
Set e = ASFAddIn.NewEngine()
Debug.Print e.Run(e.Compile("return 6 * 7;"))
```

`NewEngine` returns a new engine with `AppAccess` off, separate from the one the worksheet functions share. Importing the classes of `src/` into a workbook still works; the add-in is an alternative.

---

## LibreOffice

The extension installs a Basic library named `ASF`. Code outside it creates objects through the factories of `LOhelper`: `NewASF()`, `NewRegex()` and `NewASFMap()`.

```vb
Sub HelloASF
    GlobalScope.BasicLibraries.LoadLibrary "ASF"
    Dim engine As Object
    Set engine = ASF.LOhelper.NewASF()
    engine.Run engine.Compile("return [1,2,3,4,5].filter(fun(x){ return x > 2 }).reduce(fun(a,x){ return a+x }, 0)")
    MsgBox engine.OUTPUT_   ' => 12
End Sub
```

- **Output:** LibreOffice has no Immediate window. `print()` output is collected: `ASF.LOhelper.GetOutput()`, `ClearOutput`, and `SetOutputFile "/tmp/asf.log"` to append it to a file.
- **Documents:** with `AppAccess` on, scripts call methods and properties of UNO objects passed to them, such as `ThisComponent`.
- **Modules:** relative imports start from `engine.WorkingDir`. Files are read as UTF-8.
- **Python:** `ASF.oxt` also carries Python support, through LibreOffice's own Python. See [Python](Python.html#libreoffice).

The full details, and how to build both add-ins from the sources, are in [`excel/README.md`](https://github.com/ECP-Solutions/ASF/blob/main/excel/README.md) and [`libreoffice/README.md`](https://github.com/ECP-Solutions/ASF/blob/main/libreoffice/README.md).
