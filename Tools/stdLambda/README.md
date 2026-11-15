# stdLambda and ASF benchmark

[stdLambda](https://github.com/sancarn/stdVBA/blob/master/docs/stdLambda.md), stdVBA's lambda expressions, against ASF in Excel (VBA), on the work both can do. stdLambda is taken from stdVBA at commit `230aea3` (MIT, Copyright (c) 2020 James Warren); ASF is the checkout's `src/`.

## Workloads

Each workload measures one part of an engine. It is written once in each language and must give the same result, checked before timing (a cell that differs is recorded as `mismatch`):

| Benchmark | What it measures | Work per run | Result |
|---|---|---|---:|
| Arithmetic | Instruction throughput | `s = s + (i * 3 - i / 2) * 2`, 1000 passes | 2502500 |
| Math operations | Built-in function dispatch | `abs`, `int`, `sqr` (ASF: `Math.abs`, `Math.floor`, `Math.sqrt`), 500 passes | 292083 |
| Conditional branches | Jump performance | two `if`/`else` per pass, 1000 passes | 599 |
| For/While loops | VM dispatch efficiency | a `while` of 25 inside a `for` of 40 | 1000 |
| Function calls | Frame management | 500 calls of `add(a, b)` | 125250 |
| Fibonacci recursion | Recursive call overhead | `fib(15)`, 1973 calls | 610 |
| Array transformations | Callback and runtime overhead | `x * 2 + 1` over 300 items | 90600 (sum) |
| Object property access | Object lookup performance | `o.x * o.y` on an object of the script, 500 passes | 6000 |
| COM property access | Host integration overhead | `Count`, then `Item(i)` of a Collection of 100 | 5050 |

Each workload uses what each language has:

- **Objects**: stdLambda's `dict("x", 3, "y", 4)` (a Scripting.Dictionary, whose `o.x` reads its item), ASF's `{x: 3, y: 4}`.
- **Array transformations**: each library as it is meant to be used. stdLambda is called from VBA once per item, as `stdArray.Map` calls it, and the results go to a new array; ASF is given the array and maps it with its own function (`$1.map(fun(x) { ... })`). Both results are summed in VBA for the check.
- **COM property access**: the Collection is passed as `$1`; both read `Count` once (stdLambda's `for` evaluates its bound once, ASF stores it in a variable).

The sources are in `Benchmark.bas` (`Configure`).

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

To run one cell by hand, import the same modules into a workbook and call `LambdaBenchmark.Measure("ASF", "Fibonacci recursion", "Run")` (or `"stdLambda"`, `"Compile"`): it returns the CSV row.

## Files

- `Benchmark.bas`: workloads, checks and timing.
- `Benchmark.ps1`: the Excel runner.
- `Analyze.py`: two runs to `summary.json` and `summary.md`.
