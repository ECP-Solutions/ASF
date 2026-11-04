# Portable examples: the same scripts in Excel and LibreOffice Calc

Five ASF scripts for construction site management run unchanged in Microsoft Excel and in LibreOffice Calc. Each reads a table from a sheet and writes its result to another sheet. The business logic lives in the `.vas` files; each office suite only has a small host module that reads the table, runs the script and writes the rows back.

| Script | Reads | Writes |
|---|---|---|
| `budget.vas` | `Budget`: a bill of quantities (Chapter, Code, Description, Unit, Quantity, UnitPrice) | items, amount and share of every chapter, and the total |
| `workreports.vas` | `WorkReports`: daily work reports (Date, Code, Unit, Quantity, Crew), checked against `Budget` | one line per problem: bad codes, items not in the budget, wrong or unknown units, quantities that are not positive numbers, missing crews, dates that are not dates or fall on a Sunday |
| `progress.vas` | `Progress`: budget and measured quantities (BudgetQty, UnitPrice, PreviousQty, MonthQty) | quantity, progress and amounts to date of every item, with the items over budget, completed or without progress, and the totals |
| `price-revision.vas` | `PriceIndices`: certified amounts and labour, materials and energy indices per period | the revision coefficient Kt of a polynomial formula, the revised amount and the adjustment of every period |
| `cleanup.vas` | `Cleanup`: item lists as received from suppliers | codes, units and descriptions normalised with regular expressions, and what changed |

`scripts/lib/table.vas` is a module the scripts import: it turns the sheet table into one object per row, keyed by column name.

## How it works

```
sheet --(host reads)--> 2D array --$1, $2--> script.vas --> array of rows --(host writes)--> result sheet
```

- The hosts pass the table around cell A1 as the script's `$1` (and a second table as `$2`), with the column names in the first row and blank cells as `''`. Dates arrive as serial numbers, which are the same in both suites.
- A script returns an array of rows; the host writes it from A1 of `<sheet> result`, creating the sheet if needed. An error is written in A1 instead.
- `hosts/ExcelHost.bas` (VBA, Excel object model) and `hosts/CalcHost.bas` (LibreOffice Basic, Calc API) are the only code that knows which suite it runs in. Both have the same macros: `RunBudget`, `RunWorkReports`, `RunProgress`, `RunPriceRevision`, `RunCleanup` and `RunAll`.
- To change a rule, edit the script: the hosts stay as they are. For instance, the coefficients of the price revision formula are at the top of `price-revision.vas`.

## Running in LibreOffice Calc

1. Install `ASF.oxt` (Tools > Extensions; see [libreoffice/README.md](../../libreoffice/README.md)).
2. Open `portable-examples.ods` from this folder and allow its macros. The document already holds the five tables and the `CalcHost` module.
3. Run `RunAll` (Tools > Macros > Run Macro > portable-examples.ods > Standard > CalcHost), or one of the other macros.

## Running in Excel

1. Open `portable-examples.xlsx` from this folder (the same five tables).
2. In the VBA editor, import the class modules of `src/` (see the main [README](../../README.md#quick-start-5-minutes)) and `hosts/ExcelHost.bas`.
3. Save the workbook as `portable-examples.xlsm` in this folder, since the host looks for `scripts` next to the workbook.
4. Run `RunAll`, or one of the other macros.

## Writing portable scripts

- Read the input through the host's tables (`$1`, `$2`) and return plain values: arrays, numbers and strings. Office objects (`Range`, Calc cells) are not portable; the hosts keep them out of the scripts.
- Treat blank cells as `''`. Excel's `Empty` is converted by `ExcelHost`, because ASF reads `Empty` as `0`.
- Round with your own function when the result must match to the cent: `Math.round` rounds half to even in both suites, as VBA does. The scripts here round half away from zero.
- Keep scripts plain ASCII and write other characters as regex escapes (`²`), so that every host reads the files the same way.

## Checking the examples

```sh
python3 examples/portable/tools/expected.py   # expected results, computed independently in Python
python3 examples/portable/tools/run_lo.py     # builds the workbooks, runs the examples in LibreOffice and compares
```

`run_lo.py` builds the ASF library from `src/`, creates the five sheets from `data/*.csv`, and saves `portable-examples.ods` (with `CalcHost`) and `portable-examples.xlsx`. It then runs `RunAll` on the `.ods` and compares every result sheet with `expected/`. It needs LibreOffice and `python3-uno`.

## Files

- `scripts/`: the five scripts and `lib/table.vas`.
- `hosts/`: `ExcelHost.bas` and `CalcHost.bas`.
- `data/`: the input tables as CSV. `expected/`: the expected result tables.
- `portable-examples.ods`, `portable-examples.xlsx`: the workbooks, built by `run_lo.py`.
- `tools/`: `expected.py` and `run_lo.py`.
