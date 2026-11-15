# stdLambda and ASF benchmark

[stdLambda](https://github.com/sancarn/stdVBA/blob/master/docs/stdLambda.md), stdVBA's lambda expressions, against ASF in Excel (VBA), on direct expression evaluation: a short expression compiled once and called with arguments, as a callback is called. That is what stdLambda is made for, so the workloads are expressions, not programs, except a small loop of 5 passes that both engines run. stdLambda is taken from stdVBA at commit `230aea3` (MIT, Copyright (c) 2020 James Warren); ASF is the checkout's `src/`.

## Workloads

Each workload is written once in each language and must give the same result, checked before timing (a cell that differs is recorded as `mismatch`):

| Benchmark | What it measures | stdLambda | Arguments | Result |
|---|---|---|---|---:|
| Arithmetic | Instruction throughput | `($1 + $2) * 3 - $1 / 2` | 7, 5 | 32.5 |
| Math operations | Built-in function dispatch | `abs($1) + sqr($2) + int($1 / 3)` (ASF: `Math.abs`, `Math.sqrt`, `Math.floor`) | -7, 16 | 8 |
| Conditional branches | Jump performance | `if $1 > 5 then "big" else "small"` (ASF: `?:`) | 9 | big |
| For/While loops | VM dispatch efficiency | `let i = 0: let s = 0: do while i < 5 let i = i + 1 let s = s + i * $1 loop: s` (ASF: `while`), called 100 times from VBA | 1..100 | 75750 |
| Function calls | Frame management | `fun sq(x) x * x end: sq($1) + sq($2)` | 3, 4 | 25 |
| Fibonacci recursion | Recursive call overhead | `fib($1)` (177 calls) | 10 | 55 |
| Array transformations | Callback and runtime overhead | `$1 * 2 + 1` per item (ASF: `$1.map(fun(x) {...})`) | 1..50 | 2600 (sum) |
| Object property access | Object lookup performance | `$1.x * $1.y` on a Scripting.Dictionary (ASF: an ASF object) | `{x: 3, y: 4}` | 12 |
| COM property access | Host integration overhead | `$1.Count + $1.Item(2)` | Collection 10, 20, 30 | 23 |

The ASF sources are the same expressions with `return` (`Benchmark.bas`, `Configure`). Two workloads are called from a VBA loop:

- **For/While loops**: both engines run the same small `while` loop of 5 passes, and VBA calls it 100 times (with 1 to 100): the loop's dispatch plus the cost of each call into the engine.
- **Array transformations**: each library as it is meant to be used: stdLambda is called once per item, as `stdArray.Map` calls it, and ASF is given the array and maps it with its own function. Both results are summed in VBA for the check.

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

`Benchmark.ps1` starts a separate Excel process, imports ASF's classes (`src/*.cls`), stdLambda, stdICallable and `Benchmark.bas` into a throw-away workbook and writes `samples.csv` and `samples.json`; `-TemporaryTrustAccess` enables VBA project access for the run only. `Analyze.py` writes `summary.json` and `summary.md`: the median per call of each engine and the ASF / stdLambda ratio, with their geometric mean.

To run one cell by hand, import the same modules into a workbook and call `LambdaBenchmark.Measure("ASF", "Arithmetic", "Run")` (or `"stdLambda"`, `"Compile"`): it returns the CSV row.

## Files

- `Benchmark.bas`: workloads, checks and timing.
- `Benchmark.ps1`: the Excel runner.
- `Analyze.py`: two runs to `summary.json` and `summary.md`.
