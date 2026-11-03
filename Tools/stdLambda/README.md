# stdLambda and ASF benchmark

[stdLambda](https://github.com/sancarn/stdVBA/blob/master/docs/stdLambda.md), stdVBA's lambda expressions, against ASF on the work both can do: expressions with arguments, local variables, conditionals, loops, recursion, members of a VBA object and callbacks. stdLambda is taken from stdVBA at commit `230aea3` (MIT, Copyright (c) 2020 James Warren); ASF is the checkout's `src/`.

## Workloads and method

Each workload is written once in each language and must give the same result (checked before timing; a cell that differs is recorded as `mismatch`):

| Workload | stdLambda | ASF | Arguments | Result |
|---|---|---|---|---|
| Arithmetic expression | `($1 + $2) * 3 - $1 / 2` | `return ($1 + $2) * 3 - $1 / 2;` | 7, 5 | 32.5 |
| Predicate | `$1 mod 3 = 0 and $1 > 10` | `return $1 % 3 == 0 && $1 > 10;` | 12 | True |
| String functions | `ucase(trim($1)) & len($1)` | `return $1.trim().toUpperCase() + $1.length;` | `"  abc  "` | ABC7 |
| Conditional | `if $1 > 5 then "big" else "small"` | `return $1 > 5 ? 'big' : 'small';` | 9 | big |
| Local variables | `let a = $1 * 2: let b = a + 3: a * b` | `let a = $1 * 2; let b = a + 3; return a * b;` | 4 | 88 |
| Object members | `$1.Count + $1.Item(2)` | `return $1.count + $1.item(2);` | Collection 10, 20, 30 | 23 |
| Loop | `for i = 1 to 2000 ... next` | `for (let i = 1, i <= 2000, i += 1) {...}` | | 4000000 |
| While and branches | `do while i < 1000 if ... end ... loop` | `while (i < 1000) { if ... }` | | 2 |
| Recursion | `fun fib(n) ... end: fib($1)` | `fun fib(n) {...}; return fib($1);` | 14 | 377 |
| Callback over 300 items | `$1 * 2 + 1`, called from VBA once per item | `$1.forEach(fun(x) { s = s + x * 2 + 1; })` | 1..300 | 90600 |

Two operations per workload:

- **Compile**: a fresh object plus compiling the source: `New stdLambda` + `protInit` (bypassing `stdLambda.Create`, which returns the same object for the same text), `New ASF` + `Compile`.
- **Run**: one call of the compiled object with the arguments, as a callback is called: `lambda.Run(args)`, `engine.Run(index, args)`. For the callback workload, one call covers the 300 items: 300 `Run` calls from VBA for stdLambda (as `stdArray.Map` makes them), one `Run` with the array for ASF, which calls its own function 300 times.

Timing is Tools/stdRegex's: each cell is calibrated to at least 30 ms per batch (at most 16,384 calls), then five batches are timed; the whole experiment runs twice in separate processes, and the summary is the median of the ten per-call batch means. Engine order rotates between cells.

## Running in Excel

From the repository root, on Windows with desktop Excel, git and internet access (stdLambda and stdICallable are downloaded from stdVBA at the pinned commit):

```powershell
./Tools/stdLambda/Benchmark.ps1 -TemporaryTrustAccess -OutputDirectory ./Tools/stdLambda/benchmarks/latest
./Tools/stdLambda/Benchmark.ps1 -TemporaryTrustAccess -OutputDirectory ./Tools/stdLambda/benchmarks/latest/repeat
python ./Tools/stdLambda/Analyze.py ./Tools/stdLambda/benchmarks/latest
```

`Benchmark.ps1` starts a separate Excel process, imports ASF's classes, stdLambda, stdICallable and `Benchmark.bas` into a throw-away workbook and writes `samples.csv` and `samples.json`; `-TemporaryTrustAccess` enables VBA project access for the run only. `Analyze.py` writes `summary.json` and `summary.md`. Timing uses QueryPerformanceCounter.

## Running in LibreOffice

```sh
python3 Tools/stdLambda/BenchmarkLO.py --out=Tools/stdLambda/benchmarks/latest-lo
python3 Tools/stdLambda/BenchmarkLO.py --out=Tools/stdLambda/benchmarks/latest-lo/repeat
python3 Tools/stdLambda/Analyze.py Tools/stdLambda/benchmarks/latest-lo
```

stdLambda needs Windows (Scripting.Dictionary, VBScript.RegExp, `rtcCallByName` from VBE7.dll, `Implements`, a predeclared instance), so `lo_stdlambda.py` ports it to LibreOffice Basic:

- `LambdaDict`, a small class, stands for Scripting.Dictionary.
- The tokeniser's regexes run on `ASF_RegexEngine` (`LambdaRegExp`).
- Members are called through the edition's `CallByName` (stdLambda's Mac path, with `LOhelper.CallMember`).
- The Types move to a standard module, and arrays of Types become Variant arrays, as LibreOffice cannot `ReDim` them.
- `GoSub x: statement` lines are split, as LibreOffice resumes at the next line after `Return`.

The parser and the evaluation loop are otherwise unchanged. **stdLambda's compile times on LibreOffice are not representative**: its tokeniser tries up to 48 regexes on every token, natively compiled by VBScript in Excel but interpreted by ASF's regex engine here. Run times are comparable. LibreOffice Basic is several times slower than VBA, so compare engines within a run, not with Excel figures; its clock has 1 ms resolution.

## First run

LibreOffice 24.2.7.2, Intel Xeon @ 2.10 GHz, ASF `src/` of the perf-pass3 branch (`2e76d04`), 2 November 2026. Every cell gives the expected result in both runs. Run times, median per call:

| Workload | stdLambda | ASF | ASF / stdLambda |
|---|---|---|---|
| Arithmetic expression | 281.3 µs | 2.562 ms | 9.11× |
| Predicate | 429.7 µs | 2.187 ms | 5.09× |
| String functions | 734.4 µs | 3.062 ms | 4.17× |
| Conditional | 293.0 µs | 1.687 ms | 5.76× |
| Local variables | 304.7 µs | 2.844 ms | 9.33× |
| Object members | 773.4 µs | 3.187 ms | 4.12× |
| Loop | 732.000 ms | 4018.500 ms | 5.49× |
| While and branches | 363.000 ms | 2124.500 ms | 5.85× |
| Recursion | 277.500 ms | 3329.500 ms | 12.00× |
| Callback over 300 items | 46.500 ms | 484.500 ms | 10.42× |
| **Geometric mean** | | | **6.66×** |

stdLambda runs every workload faster, by 4 to 12 times: it compiles to a flat list of operations evaluated by one loop over a stack of Variants, with variables at fixed stack positions; ASF walks a tree of `ASF_Map` nodes and looks variables up by name in scope frames, and each `Run` sets up a program scope, which costs ASF about 2 ms per call here even for a one-line expression. Recursion is where the gap is widest (12×): a stdLambda call is a jump inside the same loop, an ASF call a new scope and a recursive evaluation. Compile times (in `summary.md`) are not comparable on LibreOffice, see above.

## Files

- `Benchmark.bas`: workloads, checks and timing.
- `Benchmark.ps1`: Excel runner. `BenchmarkLO.py`: LibreOffice runner. `lo_stdlambda.py`: stdLambda's LibreOffice port.
- `Analyze.py`: two runs to `summary.json` and `summary.md`.
- `benchmarks/2026-11-02-libreoffice/`: the first run, on LibreOffice 24.2.
