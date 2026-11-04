# Portable examples: one script, the whole office suite

Six ASF scripts run unchanged in the six office applications: Excel, Word and PowerPoint, and LibreOffice Calc, Writer and Impress. The same `budget.vas` writes a sheet in Excel or Calc, a document with headings and a table in Word or Writer, and slides in PowerPoint or Impress. Until now no macro code ran in both suites: VBA does not run in LibreOffice, and LibreOffice Basic and its UNO API do not run in Microsoft Office. With ASF the logic is written once, in `.vas` files, and each suite only needs the engine and a host of a few lines.

| Script | Does | Reads (in a spreadsheet) |
|---|---|---|
| `hello.vas` | the smallest example: a title, a section, a line and a table | nothing |
| `budget.vas` | items, amount and share of every chapter of a bill of quantities, and the total | `Budget`: Chapter, Code, Description, Unit, Quantity, UnitPrice |
| `progress.vas` | monthly progress certificate: quantity, progress and amounts to date of every item, the items over budget, completed or without progress | `Progress`: Code, Description, BudgetQty, UnitPrice, PreviousQty, MonthQty |
| `price-revision.vas` | price revision with a polynomial formula: coefficient Kt, revised amount and adjustment of every period | `PriceIndices`: Period, Certified, Labour, Materials, Energy |
| `workreports.vas` | checks of the daily work reports against the budget: bad codes, unknown items and units, quantities, crews, dates | `WorkReports` and `Budget` |
| `cleanup.vas` | codes, units and descriptions of supplier item lists normalised with regular expressions | `Cleanup`: Code, Unit, Description |

In a spreadsheet that has the sheet, a script reads it: change the numbers and run it again. Elsewhere (Word, Writer, PowerPoint, Impress, or an empty workbook) it uses the sample project of `scripts/data/project.vas`.

## How it works

```
host macro --(the open document as $1)--> script.vas --> lib/office.vas --> Excel | Word | PowerPoint | Calc | Writer | Impress
```

- The host passes the open document to the script, nothing else. `hosts/OfficeHost.bas` is one VBA module for Excel, Word and PowerPoint; `hosts/LibreOfficeHost.bas` is one Basic module for Calc, Writer and Impress. Both have the macros `RunHello`, `RunBudget`, `RunProgress`, `RunPriceRevision`, `RunWorkReports`, `RunCleanup` and `RunAll`.
- `scripts/lib/office.vas` is the only code that knows the applications. `open($1)` finds out which one holds the document and returns the same small API for the six:

  | | Excel, Calc | Word, Writer | PowerPoint, Impress |
  |---|---|---|---|
  | `doc.input(name, fallback)` | the table around A1 of sheet `name` | `fallback` | `fallback` |
  | `doc.report(title, subtitle)` | a new sheet `title` | a Heading 1 and a paragraph | a title slide |
  | `out.section(text)` | a bold row | a Heading 2 | a new slide |
  | `out.line(text)` | a row | a paragraph | a bullet |
  | `out.table(rows)` | cells (numbers stay numbers) | a table | one bullet per row |
  | `out.done()` | a line for the host to show | | |

  In Microsoft Office it calls the VBA object model (`Worksheets`, `Paragraphs`, `Slides`); in LibreOffice the UNO API (`getSheets`, text cursors, draw pages). ASF reaches both from the same script.
- `scripts/lib/table.vas` turns a table into one object per row, keyed by column name, rounds amounts and formats numbers and dates the same way everywhere.
- To change a rule, edit the script; to support another kind of output, edit `lib/office.vas`. The hosts stay as they are.

The scripts need ASF 3.3.2 or later (the classes of `src/`, or `ASF.oxt` 3.3.2).

## Running in LibreOffice (Calc, Writer, Impress)

1. Install `ASF.oxt` 3.3.2 or later (Tools > Extensions; see [libreoffice/README.md](../../libreoffice/README.md)).
2. Open `portable-examples.ods`, `portable-examples.odt` or `portable-examples.odp` from this folder and allow its macros. Each already holds the `LibreOfficeHost` module; the spreadsheet also holds the five tables.
3. Run `RunAll` (Tools > Macros > Run Macro > the document > Standard > LibreOfficeHost), or one of the other macros.

## Running in Microsoft Office (Excel, Word, PowerPoint)

1. Open `portable-examples.xlsx` (the five tables) in Excel, or a new document in Word or a new presentation in PowerPoint.
2. In the VBA editor, import the class modules of `src/` (see the main [README](../../README.md#step-1-installation)) and `hosts/OfficeHost.bas`.
3. Save it in this folder as a macro-enabled file (`.xlsm`, `.docm` or `.pptm`): the host looks for `scripts` next to the document.
4. Run `RunAll`, or one of the other macros.

## Writing portable scripts

- Talk to the document through `lib/office.vas`, not through `Range`, `Paragraphs` or UNO objects: those belong to one suite.
- Read input with `doc.input(name, fallback)`. A blank cell is `''` in Calc and an empty value in Excel; `isBlank` of `lib/table.vas` takes both as blank.
- Round with `roundTo` or `round2` of `lib/table.vas` when the result must match to the cent: `Math.round` rounds half to even, as VBA does, while these round half away from zero. Format with `show` and `dateText`, which do not depend on the locale.
- Put what several scripts share in a module and `import` it: the hosts set the scripts folder as the working directory, so `./lib/office.vas` resolves in both suites.
- A `let` inside a nested function assigns a variable of the same name in the enclosing function, as ASF functions share the scope they are written in: give inner variables their own names.
- Keep scripts plain ASCII and write other characters as escapes (`²`), so that every host reads the files the same way.

## Checking the examples

```sh
python3 examples/portable/tools/make_data.py   # scripts/data/project.vas from data/*.csv
python3 examples/portable/tools/run_lo.py      # runs every script in Calc, Writer and Impress and compares
```

`run_lo.py` builds the ASF library from `src/` and runs each script on a spreadsheet holding the tables, a text document and a presentation. It turns what the script wrote into text (sheets as tab-separated rows, paragraphs with their style, tables, slides) and compares it with `expected/<script>.<app>.txt`; `--update` rewrites those files. It also saves `portable-examples.ods`, `.odt`, `.odp` (with the host) and `.xlsx`. It needs LibreOffice and `python3-uno`. The Excel, Word and PowerPoint paths of `lib/office.vas` are not covered by it.

## Files

- `scripts/`: the six scripts; `lib/office.vas` and `lib/table.vas`; `data/project.vas`, the sample project.
- `hosts/`: `OfficeHost.bas` (Excel, Word, PowerPoint) and `LibreOfficeHost.bas` (Calc, Writer, Impress).
- `data/`: the sample tables as CSV. `expected/`: what every script writes in Calc, Writer and Impress.
- `portable-examples.ods`, `.odt`, `.odp`, `.xlsx`: the documents, built by `run_lo.py`.
- `tools/`: `make_data.py` and `run_lo.py`.
