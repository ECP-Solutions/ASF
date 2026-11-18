# ASF data pipeline for Excel

Raw orders, as several sources export them, are cleaned, checked, de-duplicated, enriched and summarized, entirely by worksheet formulas.

- Every stage is one `ASF_CALL` reading the sheet of the stage before.
- Change the raw data, the rules, the reference table or a setting, and the whole pipeline follows on recalculation.
- No macros in the workbook: the logic lives in a few `.vas` scripts in the ASF library folder.

```
Raw ──► Clean ──► Unique ──► Valid ──► Enriched ──► Sales ──► ByRegion
 │                  │                                  └────► Pivot
 │                  └──► Issues
 └──► Profile
```

| Sheet | Formula | What it gives |
|---|---|---|
| Clean | `=ASF_CALL("clean", Raw!A1:H200, Settings!B2)` | Texts trimmed; dates and numbers read from any format; codes in upper case |
| Unique | `=ASF_CALL("dedupe", Clean!A1:H200, "Order", "last")` | One row per order, the last (a correction comes after the first) |
| Issues | `=ASF_CALL("validate", Unique!A1:H200, Rules!A1:C50)` | One line per broken rule: row, order, column, value, problem |
| Valid | `=ASF_CALL("validate", Unique!A1:H200, Rules!A1:C50, "valid")` | The rows with no problem |
| Enriched | `=ASF_CALL("lookup", Valid!A1:H200, Customers!A1:D100, "Customer=CustomerID", "Name, Region, Segment")` | The customer's name, region and segment, and `Matched` |
| Sales | `=ASF_RUN(Scripts!B2, Enriched!A1:L200)` | `Amount = Qty * Price`, by a script kept in a cell |
| ByRegion | `=ASF_CALL("group", Sales!A1:M200, "Region", "sum:Amount, count:Order, avg:Amount, distinct:Customer, min:Date, max:Date")` | One row per region |
| Pivot | `=ASF_CALL("pivot", Sales!A1:M200, "Region", "Product", "sum:Amount")` | Amount by region and product, with totals |
| Profile | `=ASF_CALL("profile", Raw!A1:H200)` | What the raw data holds, column by column |

In the workbook the options come from the `Settings` sheet (`Settings!B2`…`B11`), so the pipeline is reshaped by editing cells, not formulas.

## Files

| File | |
|---|---|
| `Pipeline.xlsx` | The workbook: sample data (`Raw`, `Customers`), `Rules`, `Settings`, `Scripts` and one sheet per stage |
| `library/*.vas` | The stages, `top.vas` (the example of *Writing a stage*), and `tables.vas`, the helpers they share |
| `tools/check.py` | Runs every stage in LibreOffice and compares it with `tools/ref.py` |
| `tools/ref.py` | The same stages written again in Python, from their descriptions |
| `tools/make_workbook.py` | Writes `Pipeline.xlsx` from `tools/data.py` |
| `tools/timing.py` | Times every stage in LibreOffice on 1,000 orders |

## Setting up

1. **Load the ASF add-in** (`excel/ASF.xlam`; [excel/README.md](../../excel/README.md) says how to build and install it). It must be built from `addin-engine` at commit `f4b4283` or later. That build has engine fixes the pipeline relies on:
   - an empty cell inside a range is `undefined`, not `0`;
   - `for…of` over an empty list doesn't run;
   - `sort` is stable, which `top.vas` relies on.
2. **Copy `library/*.vas` into the ASF library folder.** The ASF tab's **Library folder** button opens it.
3. **Open `Pipeline.xlsx`.** It recalculates on opening.

Each stage is an array formula over a fixed range, such as `Clean!A1:H200`. That works in Excel 2016 and in Microsoft 365 alike: unused cells stay empty, and the next stage skips empty rows and columns. Rows added to `Raw`, up to row 200, flow through every stage.

In Microsoft 365 you can instead write each stage as one formula in `A1` that spills, and point the next stage at it with `#`: `=ASF_CALL("dedupe", Clean!A1#, "Order", "last")`.

## The stages

Every stage takes a table with a header row. Columns are named by their headers, without regard to case.

**`clean` (table, spec):** the spec says how to read some columns, as `Column:kind, …`:

| Kind | Reads |
|---|---|
| `number` | `$1,250.00`, `1.250,50`, `€ 9,99`, `(3.00)` (negative), `15%`, `-12` |
| `date`, `date-dmy`, `date-mdy` | Excel dates; `yyyy-mm-dd` text; `d/m/y` or `m/d/y` text with `/`, `-` or `.`, and two-digit years as 20yy. Impossible dates such as `31/02/2026` are left as they were |
| `lower`, `upper`, `title` | The text in that case |
| `text` | Numbers as text too (codes, phone numbers) |

Also:
- Every text is trimmed, and runs of spaces become one.
- Empty rows go.
- A value a kind can't read stays as it was, so `validate` can report it.
- In a number, one separator followed by groups of three digits (`1,250`, `1.234.567`) is a thousands separator; any other is the decimal point.

**`dedupe` (table, keys, [first | last]):** one row per key, the first or the last of each, kept in its place. Keys are compared as trimmed text, without regard to case.

**`validate` (table, rules, [issues | valid | flagged]):** the rules are a table of `Column | Rule | Value`:

| Rule | The cell… |
|---|---|
| `required` | is not empty |
| `number`, `date` | holds a number (dates are numbers after `clean`) |
| `min`, `max` | is a number not below / above Value |
| `pattern` | matches the regular expression in Value (JavaScript syntax) |
| `oneOf` | is one of Value's comma-separated items, without regard to case |
| `unique` | holds a value no other row has |

- Every rule but `required` skips empty cells.
- `issues` lists each problem once per cell. `Row` counts the header as 1, and the second column is the table's first column, usually the id.
- `valid` gives the rows without problems.
- `flagged` gives the whole table, with an `Issues` column.

**`lookup` (table, reference, key, [columns]):** adds columns of the reference table, like `XLOOKUP` for whole columns.
- The key is `Column`, when both tables name it the same, or `Column=ReferenceColumn`.
- Unmatched rows get empty cells, and the `Matched` column (`TRUE`/`FALSE`) shows which they are. In the sample, the unknown customer `C099` ends up under the region `(blank)` in `ByRegion`.

**`group` (table, by, measures):** one row per group, sorted.
- `by` lists the columns to group by. Leave it empty for one row for the whole table.
- Measures: `sum`, `avg`, `min`, `max` (of the numbers), `count` (non-empty cells; `count:*` counts rows) and `distinct`.

**`pivot` (table, rows, columns, [measure]):** a cross-tab with a total row and column. The measure is `sum:Column`, `avg:Column` or `count`, the default.

**`profile` (table):** per column:
- filled, empty and distinct cells;
- how many are numbers, and how many are text that reads as a number;
- the smallest and largest number;
- the most frequent value and how often it appears.

Run it on raw data to see what `clean` has to fix. In the sample, 24 of 25 prices are text.

When a stage gets a wrong argument, its cells show `#VALUE!`, and **Why this error?** on the ASF tab names the problem, for example `No column 'Quantity' in the table` or `Unknown measure 'total'`. A stage that reads a failed stage says so, rather than working on `#VALUE!` cells.

## Writing a stage

A stage is a `.vas` file in the library folder. `tables.vas` gives it the table helpers. This one, `top.vas`, is in the library:

```javascript
// @asf name: Top N
// top.vas: the n rows with the largest value of a column. =ASF_CALL("top", Sales!A1:M200, "Amount", 5)
import { trimTable, columnIndex, column } from './tables.vas';

let t = trimTable($1);
let c = column(columnIndex(t), $2);
let rows = t.slice(2);
// largest first; equal values keep their order (the sort is stable)
rows.sort(fun(a, b) { return b[c] - a[c]; });
let out = [t[1]];
for (let i = 1, i <= rows.length && i <= $3, i += 1) { out.push(rows[i]); };
return out;
```

Things to know when writing ASF for tables:

| What | How it works |
|---|---|
| Indexes | Arrays count from 1 (`t[1]` is the header row); strings from 0, as in JavaScript (`s.charAt(0)`) |
| Sorting | Stable, as in JavaScript: equal items keep their order |
| Empty values | `0 == undefined` and `'' == undefined` are true, so test empty cells with `typeof v == 'undefined'`, as `isBlank` does |
| Values, not references | Arrays are copied when assigned; objects (`{}`) are shared. To change an array kept in an object, `set` it again |

## Size and speed

ASF runs as VBA, so a stage costs roughly a fixed amount per cell. For the whole chain on 1,000 orders, measured in LibreOffice, whose Basic is much slower than VBA:

| Stage | Time in LibreOffice |
|---|---|
| clean | 92 s |
| dedupe | 31 s |
| validate | 73 s |
| lookup | 11 s |
| amount (ASF_RUN) | 10 s |
| group | 38 s |
| pivot | 24 s |
| profile | 130 s |
| **all** | **409 s** |

For a rough Excel figure, divide by about 15 to 20, the ratio measured between the two for ASF runs. That means a few seconds per stage at a thousand rows, longer for `clean` and `profile`, so:
- keep the ranges near the data's size;
- switch to manual calculation (Formulas > Calculation Options) while pasting large data;
- keep `profile` for checking new data rather than in every recalculation.

`dedupe`, `group`, `profile` and `validate`'s `unique` rule keep values in ASF objects, whose lookups scan their keys one by one. These stages slow down with many distinct keys.

## Checking it

From the repository:

```
python3 examples/excel-pipeline/tools/check.py
```

This needs LibreOffice and `python3-uno`. It runs the workbook's chain through the add-in's own modules, each stage fed with what the stage before gave. Each result is compared with `tools/ref.py`, which runs its own chain. It also checks:
- the number and date readers, on their own;
- a table padded the way an array formula pads it;
- that wrong arguments give `#VALUE!` with a message naming the problem.

After changing `tools/data.py`, `python3 examples/excel-pipeline/tools/make_workbook.py` writes the workbook again, and `tools/timing.py` times the stages.
