# Regex engine benchmarks

The regex benchmark of [stdVBA](https://github.com/sancarn/stdVBA) (`Tools/stdRegex`, branch `codex/issue-141-regex-analysis`, commit `0bd56cf`), which compares stdRegex's engines with ASF's for [stdVBA#141](https://github.com/sancarn/stdVBA/issues/141), reproduced here and extended with more ASF versions. Workloads, checks, timing and statistics are unchanged; stdVBA is MIT-licensed (Copyright (c) 2021 James Warren).

| Engine | Source | Runs in |
|---|---|---|
| Old stdRegex | stdVBA `abf14dd`, VBScript `RegExp` | Excel (Windows) |
| Sihlfall | stdVBA `04e7d04`, stdRegex on [sihlfall/vba-regex](https://github.com/sihlfall/vba-regex) `1619cd7` | Excel (Windows) |
| ASF 290ad3e | `ASF_RegexEngine` at `290ad3e`, the version stdVBA measured | Excel, LibreOffice |
| ASF v3.2.0 | `ASF_RegexEngine` at tag `v3.2.0` | Excel, LibreOffice |
| ASF | `src/ASF_RegexEngine.cls` of the checkout | Excel, LibreOffice |

The stdRegex engines need Windows: the old one wraps VBScript, the new one returns `Scripting.Dictionary` results. The older ASF versions are read from git history and renamed (`ASF_RegexEngine_290ad3e`, `ASF_RegexEngine_v320`) so all versions coexist, as stdVBA does with the old stdRegex.

## Workloads and method

Eight workloads (literal scan, short validation, capture extraction, case-insensitive alternation, multiline anchors, no-match scan, word boundaries, backreference), each measured for four operations: fresh object + initialization + first `Test`, reused `Test`, `MatchAll` (`ExecAll` for ASF) and global `Replace`. Before timing, every cell checks the expected Boolean, every numbered capture and the match count, or the complete replacement string; a cell that fails is recorded as `mismatch` and left out of comparisons. Each cell is calibrated to at least 30 ms per batch (at most 16,384 calls), then five batches are timed. The whole experiment runs twice in separate processes; the summary is the median of the ten per-call batch means. Engine order rotates between cells. Step limits are raised to 2,147,483,647.

## Running in Excel

From the repository root, on Windows with desktop Excel, git and internet access (the stdRegex classes are downloaded from stdVBA at the pinned commits):

```powershell
./Tools/stdRegex/Benchmark.ps1 -TemporaryTrustAccess -OutputDirectory ./Tools/stdRegex/benchmarks/latest
./Tools/stdRegex/Benchmark.ps1 -TemporaryTrustAccess -OutputDirectory ./Tools/stdRegex/benchmarks/latest/repeat
python ./Tools/stdRegex/Analyze.py ./Tools/stdRegex/benchmarks/latest
python ./Tools/stdRegex/BuildReport.py ./Tools/stdRegex/benchmarks/latest
```

`Benchmark.ps1` starts a separate Excel process, imports the five engine classes and `Benchmark.bas` into a throw-away workbook and writes `samples.csv` and `samples.json`. `-TemporaryTrustAccess` enables VBA project access for the run and restores the previous setting afterwards; debugger settings are restored too. `-Engines` measures a subset, for example `-Engines 'ASF','Sihlfall'`. Timing uses QueryPerformanceCounter.

## Running in LibreOffice

From the repository root, with LibreOffice, `python3-uno` and git:

```sh
python3 Tools/stdRegex/BenchmarkLO.py --out=Tools/stdRegex/benchmarks/latest-lo
python3 Tools/stdRegex/BenchmarkLO.py --out=Tools/stdRegex/benchmarks/latest-lo/repeat
python3 Tools/stdRegex/Analyze.py Tools/stdRegex/benchmarks/latest-lo
python3 Tools/stdRegex/BuildReport.py Tools/stdRegex/benchmarks/latest-lo
```

`BenchmarkLO.py` builds the LibreOffice edition (see `libreoffice/README.md`) with the three ASF engines and runs `Benchmark.bas` headless; `--engines` measures a subset. LibreOffice Basic is several times slower than VBA, so compare engines within a run, not with Excel figures. Its clock has 1 ms resolution, about 3% of a 30 ms batch.

## Files

- `Benchmark.bas`: workloads, checks and timing (stdVBA's, with the ASF versions added).
- `Benchmark.ps1`: Excel runner. `BenchmarkLO.py`: LibreOffice runner.
- `Analyze.py`: two runs to `summary.json`. `BuildReport.py`: `summary.json` to a self-contained `report.html`, with findings computed from the data.
- `benchmarks/2026-10-27-stdvba-excel/`: stdVBA's recorded Excel run (old stdRegex, Sihlfall, ASF `290ad3e`), copied unchanged with its report.
- `benchmarks/2026-10-30-libreoffice/`: the first run here, on LibreOffice 24.2 with ASF `290ad3e`, v3.2.0 and the regex-phase2 engine.
- `benchmarks/2026-10-30-stdvba-excel/`: Excel 16.0 run with all five implementations; ASF is the v3.2.1 engine.
