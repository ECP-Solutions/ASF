# stdLambda and ASF benchmark

[stdLambda](https://github.com/sancarn/stdVBA/blob/master/docs/stdLambda.md), stdVBA's lambda expressions, against ASF in Excel (VBA). stdLambda is taken from stdVBA at commit `230aea3` (MIT, Copyright (c) 2020 James Warren); ASF is the checkout's `src/`.

## Workloads

21 workloads in seven categories. Each is written once in each language and must give the same result, checked before timing (a cell that differs is recorded as `mismatch`). The sources are in `Benchmark.bas` (`Configure`, and `NumberTheory...` / `TextReport...` for the long programs).

| Category | Workload | stdLambda | Arguments | Result |
|---|---|---|---|---:|
| Simple expressions | Sum and product | `($1 + $2) * 3 - $1 / 2` | 7, 5 | 32.5 |
| | Power and modulo | `$1 ^ 2 + $2 mod 4 - ($1 - $2) / 2` (ASF: `%`) | 7, 5 | 49 |
| | Comparison and logic | `if $1 > 5 and $2 <= 5 or $1 = $2 then "yes" else "no"` (ASF: `&&`, `\|\|`, `?:`) | 7, 5 | yes |
| | String concatenation | `$1 & "-" & $2 & "!"` (ASF: `+`) | "abc", 7 | abc-7! |
| Math functions | Rounding and roots | `abs($1) + sqr($2) + int($1 / 3)` (ASF: `Math.abs`, `Math.sqrt`, `Math.floor`) | -7, 16 | 8 |
| | Trigonometry and logarithms | `sin($1) ^ 2 + cos($1) ^ 2 + exp(log($2)) + atn(1) * 4` | 7, 5 | 9.14159... |
| | String functions | `ucase(left($1, 3)) & len($1) & mid($1, 5, 3)` (ASF: `slice`, `toUpperCase`, `length`, `substring`) | "benchmark" | BEN9hma |
| Script functions | One function | `fun sq(x) x * x end: sq($1) + sq($2)` | 3, 4 | 25 |
| | Nested calls | `add`, `mul` and `poly(x) = add(mul(x, x), mul(3, x))`: `poly($1) + poly($2)` | 3, 4 | 46 |
| | Recursion | `fib($1)` (177 calls) | 10 | 55 |
| COM objects | Range properties | `$1.Value * 2 + $1.Row + $1.Column` | cell B3 (21) | 47 |
| | Range methods | `$1.Offset(1, 2).Value + $1.Worksheet.Range("C4").Value` | cell B3 | 70 |
| | FileSystemObject methods | `len($1.GetBaseName("report.txt")) + len($1.GetExtensionName(...)) + len($1.BuildPath("a", "b"))` | Scripting.FileSystemObject | 12 |
| COM collections | Collection items | `$1.Count + $1.Item(2)` | Collection 10, 20, 30 | 23 |
| | Dictionary lookups | `$1.count + $1.item("b") + (if $1.exists("z") then 100 else 0)` | Scripting.Dictionary a, b, c | 5 |
| | Collection loop | `let s = 0: for i = 1 to $1.count let s = s + $1.item(i) next: s` | Collection 10..100 | 550 |
| Short programs | Variables and branch | `let a = $1 * 2: let b = a + $2: if b > 10 then b - 10 else b end` | 7, 5 | 9 |
| | Loop of 5, called 100 times | `let s = 0: for i = 1 to 5 let s = s + i * $1 next: s`, called from VBA with 1 to 100 | 1..100 | 75750 |
| | String building | `let s = "": for i = 1 to 5 let s = s & i next: s & $1` | "!" | 12345! |
| Long programs | Number theory | 45 lines: `isprime`, `gcd` and `collatz` with `do while` loops and branches, summed over 1 to `$1` | 30 | 10097441 |
| | Text report | 36 lines: `grade` (chained `if`), `pad` (a loop), `score`; a report string built over 1 to `$1` | 20, 37 | 120\|939\|94\|..1:C;... |

The categories measure:

- **Simple expressions**: operators only, the instruction throughput of each engine.
- **Math functions**: the functions of each language (stdLambda's `abs`, `sqr`, `sin`, `ucase`, `mid`...; ASF's `Math` and string methods).
- **Script functions**: functions declared in the script: calls, nested calls and recursion.
- **COM objects**: properties and methods of host objects that are not collections: Excel ranges (written in the first sheet of the benchmark's workbook) and `Scripting.FileSystemObject` (its path functions, which do not touch the disk).
- **COM collections**: `Collection` and `Scripting.Dictionary`, by key and in a loop over the items.
- **Short programs** and **long programs**: variables, loops, branches and functions together; the long programs have 34 to 45 lines.

### stdLambda's dialect

The stdLambda sources keep to what its parser at `230aea3` handles; each of these was checked by running it:

- A `for` loop's variable is declared before the loop (`let k = 0`): with a new variable, a body that sets two or more variables gives wrong results.
- Every `if` has its own `end`: in `if ... else if ... else ... end` inside a function, the inner `if` takes the single `end` and the outer `else` takes the function's, so the chain is written `... else ... end end end`.
- A function sees only its parameters (not the variables of the program), and calls only functions declared before it.
- `true` and `false` are written in lower case (`True` is read as false), and the members of a `Scripting.Dictionary` are written in lower case (`count`, `item`, `exists`): other names read keys of the dictionary.
- Keywords and functions are otherwise case-insensitive; a multi-line program is its lines joined by spaces, as `stdLambda.CreateMultiline` joins them, and its value is the value of its last statement.

## Method

Two operations per workload:

- **Compile**: a fresh object plus compiling the source: `New stdLambda` + `protInit` (bypassing `stdLambda.Create`, which returns the same object for the same text), `New ASF` + `Compile`.
- **Run**: one call of the compiled object with the arguments, as a callback is called: `lambda.Run(args)`, `engine.Run(index, args)`.

Timing is Tools/stdRegex's: each cell is calibrated to at least 30 ms per batch (at most 16,384 calls), then five batches are timed with QueryPerformanceCounter; the whole experiment runs twice in separate Excel processes, and the summary is the median of the ten per-call batch means. Engine order rotates between cells.

## Running it

From the repository root, on Windows with desktop Excel, git and internet access (stdLambda and stdICallable are downloaded from stdVBA at the pinned commit):

```powershell
./Tools/stdLambda/Benchmark.ps1 -TemporaryTrustAccess -OutputDirectory ./Tools/stdLambda/benchmarks/latest
./Tools/stdLambda/Benchmark.ps1 -TemporaryTrustAccess -OutputDirectory ./Tools/stdLambda/benchmarks/latest/repeat
python ./Tools/stdLambda/Analyze.py ./Tools/stdLambda/benchmarks/latest
```

`Benchmark.ps1` starts a separate Excel process, imports ASF's classes (`src/*.cls`), stdLambda, stdICallable and `Benchmark.bas` into a throw-away workbook and writes `samples.csv` and `samples.json`; `-TemporaryTrustAccess` enables VBA project access for the run only. `Analyze.py` writes `summary.json` and `summary.md`: the median per call of each engine and the ASF / stdLambda ratio, with its geometric mean per category and over all workloads.

To run one cell by hand, import the same modules into a workbook and call `LambdaBenchmark.Measure("ASF", "Sum and product", "Run")` (or `"stdLambda"`, `"Compile"`): it returns the CSV row.

## Files

- `Benchmark.bas`: workloads (`Workloads` lists them by category), checks and timing.
- `Benchmark.ps1`: the Excel runner.
- `Analyze.py`: two runs to `summary.json` and `summary.md`.
