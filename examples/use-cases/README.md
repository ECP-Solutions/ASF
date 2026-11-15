# Use cases

Four ways to put ASF to work, each with code that runs and is checked:

| # | Use case | For | What ASF brings | Where |
|---|---|---|---|---|
| 1 | [Business rules users edit](#1-business-rules-users-edit-without-touching-vba) | finance, pricing, quality teams | rules in cells, checked and compiled when they run; no VBA editing, no access to the VBA project | [`rules/`](rules) |
| 3 | [Regular expressions everywhere](#3-regular-expressions-that-work-everywhere) | Mac users, LibreOffice users, anyone leaving VBScript | a regex engine written in VBA, with JavaScript's syntax, as worksheet functions or in scripts | [`regex/`](regex) |
| 5 | [Data pipelines over sheet data](#5-data-pipelines-over-sheet-data) | analysts who summarise tables | one read of the sheet, grouping and sorting in ASF, one write | [`pipeline/`](pipeline) |
| 7 | [JavaScript-style scripting where Office Scripts does not reach](#7-javascript-style-scripting-where-office-scripts-does-not-reach) | people who know JavaScript and automate Office | modules, classes and `extends`, regex, templates; one script for six applications | [`../portable`](../portable) |

## Before you start

- **Modules.** Import ASF's class modules (`src/*.cls`) into the workbook, and the modules each use case names. Use cases 1 and 5 also use [`ASFSheets.bas`](ASFSheets.bas). In LibreOffice, install `ASF.oxt`.
- **Tables in, tables out.** `Range.Value2` goes to a script as it is: ASF reads a 2D array as an array of rows, from 1 like the sheet (`$1[1]` is the header row, `$1[i][c]` a cell). A script returns an array of rows; `ToGrid` of `ASFSheets.bas` turns it back into a 2D array for `Range.Value2`, and `WriteRows` writes it in one assignment.
- **Errors.** `Compile` raises an error for a source that does not compile, with a message that says why. `Run` and `Execute` do not raise: they return `Empty` and leave the message (or the thrown value) in `engine.LastError`.
- **Access to Office.** `AppAccess` is `False` unless the host sets it: a script then sees only the values it is given (`$1`, `$2`...), not the objects of the application.

---

## 1. Business rules users edit without touching VBA

Discounts, eligibility, validations and scoring change every month. When they live in VBA, every change goes through whoever can edit the macros.

Here the rules live in a sheet, one per row, written as ASF expressions over the fields of an order. The host builds a small program from them, checks each rule on its own so that a mistake names its rule, compiles everything once and runs it once over the whole table.

```
Rules sheet ──► RulesHost.BuildConditions ──► let conditions = [fun(o) { ... }, ...];
                                                   + discounts.vas
Orders sheet (Value2) ──► engine.Run(idx, orders, rules) ──► Discounts sheet (one write)
```

**Files.** [`rules/discounts.xlsx`](rules/discounts.xlsx) (sheets Rules, Orders, Discounts), [`rules/discounts.vas`](rules/discounts.vas), [`rules/RulesHost.bas`](rules/RulesHost.bas). Save the workbook as `.xlsm` next to `discounts.vas`, import the ASF classes, `ASFSheets.bas` and `RulesHost.bas`, and run `ApplyDiscountRules`.

**Rules**: what the users edit.

| Name | When | Discount |
|---|---|---:|
| Volume | `o.qty >= 100` | 0.08 |
| Loyal EU customer | `o.region == 'EU' && o.years >= 3` | 0.05 |
| Clearance | `o.sku.startsWith('CL-')` | 0.15 |
| Big ticket | `o.qty * o.price > 5000` | 0.10 |

**Orders**: the column names are the names the rules use (`o.qty` reads the column `qty`).

| order | region | years | sku | qty | price |
|---|---|---:|---|---:|---:|
| A-101 | EU | 4 | ST-200 | 20 | 35 |
| A-102 | US | 1 | CL-310 | 5 | 12 |
| A-103 | EU | 1 | ST-200 | 150 | 35 |
| A-104 | US | 6 | ST-900 | 60 | 120 |
| A-105 | EU | 3 | CL-110 | 200 | 3 |

**Discounts**: the best discount that applies to each order.

| Order | Discount | Rule |
|---|---:|---|
| A-101 | 0.05 | Loyal EU customer |
| A-102 | 0.15 | Clearance |
| A-103 | 0.10 | Big ticket |
| A-104 | 0.10 | Big ticket |
| A-105 | 0.15 | Clearance |

A mistake in a rule stops the run with a message that names the rule, before any order is priced:

| The user writes | The message |
|---|---|
| `o.qty >= (100` | Rule 'Volume': Unbalanced opening or closing parentheses. |
| `o.qty >=` | Rule 'Volume': Expected an expression after >= |
| `o.skuu.startsWith('CL-')` | Stopped: Rule 'Volume' uses o.skuu, but Orders has no column 'skuu' |
| `o.qtty >= 100` | Stopped: Rule 'Volume' uses o.qtty, but Orders has no column 'qtty' |

**Design notes.**

- **Rules are code, the rest is data.** Only the When column becomes code; names and discounts are read from `$2`, so a name with quotes or a discount in any locale cannot break the program.
- **Two checks.** `CheckRules` compiles each rule alone, so a syntax error names its rule. `discounts.vas` then checks, with ASF's own regex, that every `o.<column>` exists: a misspelt column would otherwise make a rule quietly false, as in JavaScript (`undefined >= 100` is false).
- **One compile, one run.** The rules are compiled once per click and run over the whole table in one `Run`: the cost is one call into ASF, not one per order.
- **Safe by default.** The engine keeps `AppAccess = False`: a rule reads the order it is given, not the workbook or the file system.

---

## 3. Regular expressions that work everywhere

VBA has no regular expressions of its own. `VBScript.RegExp` exists only on Windows (not in Excel for Mac), speaks VBScript's dialect (no lookbehind, no named groups) and comes with VBScript, which Microsoft is phasing out. LibreOffice Basic has none.

`ASF_RegexEngine` is a regex engine written in VBA with JavaScript's syntax, checked by 182 tests and against Node.js on a corpus of about 15,700 cases. It works without a script, as a class.

**Files.** [`regex/RegexFunctions.bas`](regex/RegexFunctions.bas) (Excel, Windows and Mac: import it with `ASF_RegexEngine.cls`), [`regex/RegexFunctionsCalc.bas`](regex/RegexFunctionsCalc.bas) (LibreOffice Calc: add it to the Standard library of the spreadsheet or of My Macros, with `ASF.oxt` installed), [`regex/regex-functions.xlsx`](regex/regex-functions.xlsx) (the examples below as formulas; after importing the module, recalculate with Ctrl+Alt+F9).

| Formula | Result |
|---|---|
| `=RX_TEST("ES-0042", "^[A-Z]{2}-\d{4}$")` | TRUE |
| `=RX_TEST("es-0042", "^[A-Z]{2}-\d{4}$")` | FALSE |
| `=RX_TEST("es-0042", "^[A-Z]{2}-\d{4}$", TRUE)` (ignore case) | TRUE |
| `=RX_EXTRACT("Pallet: 12 boxes, 340 kg", "(\d+)\s*kg", 1)` (a group) | 340 |
| `=RX_EXTRACT("Pallet: 12 boxes, 340 kg", "\d+")` | 12 |
| `=RX_EXTRACT("no numbers", "\d+")` | (empty) |
| `=RX_EXTRACT("Order #A-77 shipped", "(?<=#)[A-Z]-\d+")` (lookbehind) | A-77 |
| `=RX_EXTRACT("2026-11-15", "(?<year>\d{4})-(?<month>\d{2})", 2)` (named groups) | 11 |
| `=RX_REPLACE("  too    many   spaces ", "\s+", " ")` | ` too many spaces ` |
| `=RX_REPLACE("Smith, John", "^(\w+),\s*(\w+)$", "$2 $1")` | John Smith |
| `=RX_COUNT("one two  three", "\b\w+\b")` | 3 |
| `=RX_TEST("x", "(unclosed")` | #VALUE! |

The results are those of JavaScript (Node.js) for the same patterns. In Calc, arguments are separated by `;`, and a bad pattern gives the text `#VALUE!`: a Basic function cannot return Calc's error values.

The engine keeps the patterns it has compiled, so a column of formulas with the same pattern compiles it once.

**Cleaning whole tables.** A script uses the same engine through `regex(...)` and the string methods (`replace`, `match`, `split`). [`../portable/scripts/cleanup.vas`](../portable/scripts/cleanup.vas) normalises the codes, units and descriptions of supplier item lists, and runs unchanged in Excel, Word, PowerPoint, Calc, Writer and Impress.

---

## 5. Data pipelines over sheet data

Summaries written in VBA mix two kinds of code: the calls into Excel (slow, one per cell when written naively) and the logic (grouping, totals, ranking). The pattern here touches Excel twice: one read of `Value2`, one write of the result. In between, ASF works on plain values, with objects as dictionaries, `map`, `sort` and closures.

**Files.** [`pipeline/sales.xlsx`](pipeline/sales.xlsx) (sheets Sales, Summary), [`pipeline/sales-summary.vas`](pipeline/sales-summary.vas), [`pipeline/SalesHost.bas`](pipeline/SalesHost.bas). Save the workbook as `.xlsm` next to `sales-summary.vas`, import the ASF classes, `ASFSheets.bas` and `SalesHost.bas`, and run `SummarizeSales`.

For the 36 sample rows (three regions, four products, three months):

| Region | Units | Revenue | Share % | Top product |
|---|---:|---:|---:|---|
| South | 263 | 573 | 36.77 | Washers |
| East | 255 | 526.5 | 33.78 | Washers |
| North | 248 | 459 | 29.45 | Bolts |

| Month | Revenue | Change % |
|---|---:|---:|
| 2026-01 | 576.5 | |
| 2026-02 | 456.5 | -20.82 |
| 2026-03 | 525.5 | 15.12 |

**Notes.**

- **Columns by name.** The script finds its columns by their header, so the sheet's columns can move.
- **Month as text.** Excel turns `2026-01` typed in a cell into a date: the Month column of `sales.xlsx` is formatted as text.
- **Measure on your data.** Time grows linearly with the rows; put `Timer` around `Run` to see the cost on your tables. LibreOffice Basic is much slower than VBA, so the pattern suits thousands of rows there rather than hundreds of thousands.

---

## 7. JavaScript-style scripting where Office Scripts does not reach

Office Scripts gives TypeScript to Excel on the web. On the desktop, in Word and PowerPoint, and in LibreOffice, people who know JavaScript are left with VBA or Basic. ASF gives them a language close to JavaScript in all of them: modules (`import`/`export`), classes with `extends` across files, closures, template literals, regex, `try`/`catch`.

**Files.** [`../portable/scripts/quality-report.vas`](../portable/scripts/quality-report.vas) and [`../portable/scripts/lib/checks.vas`](../portable/scripts/lib/checks.vas), with the portable examples: run `RunQualityReport` of `OfficeHost.bas` (Excel, Word, PowerPoint) or `LibreOfficeHost.bas` (Calc, Writer, Impress), as [their README](../portable/README.md) explains.

`lib/checks.vas` holds four classes: `Check` runs a check over every record, and `RequiredCheck`, `RangeCheck` and `PatternCheck` extend it and only say what failing means. `quality-report.vas` imports them, checks the readings of a production line (the sheet Readings, or the sample readings elsewhere) and writes the report through `lib/office.vas`, which knows the six applications:

| Check | Failed | Rows |
|---|---:|---|
| Batch code is not B0000-X | 2 | 5, 8 |
| Temperature outside 180 to 220 | 2 | 4, 8 |
| Pressure outside 1.5 to 2.5 | 1 | 6 |
| Inspector is missing | 1 | 6 |

A sheet in Excel or Calc, a document in Word or Writer, slides in PowerPoint or Impress: the same script.

---

## How the use cases are checked

```bash
python3 examples/use-cases/tools/run_lo.py          # the scripts and modules in LibreOffice
python3 examples/use-cases/tools/make_workbooks.py  # writes the sample workbooks (needs openpyxl)
python3 examples/portable/tools/run_lo.py           # use case 7, with the other portable examples
```

`tools/run_lo.py` builds the ASF library from `src/` and runs, in LibreOffice: the rules of `discounts.xlsx` through `RulesHost`'s `BuildConditions` and `CheckRules` and `discounts.vas`, and the four mistakes above; `sales-summary.vas` on `sales.xlsx`; the functions of `RegexFunctions.bas` called from Basic, and those of `RegexFunctionsCalc.bas` in the cells of a spreadsheet. It compares the results with [`tools/data.py`](tools/data.py), which holds the sample tables, the expected discounts and regex results, and its own calculation of the sales summary. The macros that only Excel runs (`ApplyDiscountRules`, `SummarizeSales`, `WriteRows`) are not run by it.

## Files

- `ASFSheets.bas`: `ToGrid` and `WriteRows`, shared by use cases 1 and 5.
- `rules/`: `discounts.xlsx`, `discounts.vas`, `RulesHost.bas`.
- `regex/`: `RegexFunctions.bas` (Excel), `RegexFunctionsCalc.bas` (Calc), `regex-functions.xlsx`.
- `pipeline/`: `sales.xlsx`, `sales-summary.vas`, `SalesHost.bas`.
- `tools/`: `data.py`, `make_workbooks.py`, `run_lo.py`.
