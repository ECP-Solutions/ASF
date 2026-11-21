# ASF for LibreOffice

The LibreOffice edition of ASF: the same engine, language, regex engine and module system, packaged as a LibreOffice Basic library named `ASF` and shipped as the `ASF.oxt` extension.

It is generated from the VBA sources in [`src/`](../src), which stay the single source of truth: [`tools/build.py`](tools/build.py) rewrites each class for LibreOffice Basic (`Option VBASupport 1`). Every rewrite works around one documented difference between VBA and LibreOffice Basic, so a fix or feature in `src/` reaches both editions.

## Requirements

- LibreOffice 7.5 or later (tested on 24.2).
- For `@(...)` VBA expressions only: the [VBA Expressions extension](https://extensions.libreoffice.org/en/extensions/show/70059) 1.2.1 or later (`VBAExpressionsLib`). ASF works without it until a script evaluates an `@(...)` expression.

## Installation

[`ASF.oxt`](ASF.oxt), in this folder, is the extension built from `src/`. Install it with *Tools > Extension Manager > Add...*, or from a terminal:

```sh
unopkg add ASF.oxt
```

Install the VBA Expressions extension the same way if your scripts use `@(...)`.

## Usage

```vb
Sub HelloASF
    GlobalScope.BasicLibraries.LoadLibrary "ASF"
    Dim engine As Object
    Set engine = ASF.LOhelper.NewASF()
    engine.Run engine.Compile("return [1,2,3,4,5].filter(fun(x){ return x > 2 }).reduce(fun(a,x){ return a+x }, 0)")
    MsgBox engine.OUTPUT_   ' => 12
End Sub
```

Code outside the `ASF` library creates objects through the factories in `LOhelper`:

| Factory | Returns |
|---|---|
| `NewASF()` | an `ASF` engine |
| `NewRegex()` | an `ASF_RegexEngine` |
| `NewASFMap()` | an `ASF_Map` |

### Output

LibreOffice Basic has no Immediate window. Text that the VBA edition sends there with `Debug.Print`, such as `print()` output and the runtime log of `verbose` engines, is collected instead:

```vb
ASF.LOhelper.SetOutputFile "/tmp/asf.log"   ' also append every line to a file ("" stops it)
engine.Run engine.Compile("print('hello');")
MsgBox ASF.LOhelper.GetOutput()             ' everything printed so far
ASF.LOhelper.ClearOutput
```

### Documents and UNO objects

With `AppAccess` on, scripts call methods and properties of UNO objects passed to them, as the VBA edition does with the Office object model:

```vb
engine.AppAccess = True
engine.Run engine.Compile( _
    "let cell = $1.getSheets().getByIndex(0).getCellByPosition(0, 0);" & _
    "cell.String = 'Hello from ASF'; return typeof cell;"), ThisComponent
' => object: <ScCellObj>
```

UNO members are found case-insensitively; `typeof` reports a UNO object's implementation name.

### Modules

`import`/`export` work as in the VBA edition. Relative imports start from `engine.WorkingDir` (or `scwd()` inside a script), and inside a module from the module's folder. Script and module files are read as UTF-8.

### VBA expressions

`@(...)` expressions run on `VBAExpressionsLib`. User-defined functions are declared with `DeclareUDF` and live in that library's `UDFunctions` module, as described in its documentation; the functions in this repository's `src/UDFunctions.cls` are not part of the LibreOffice edition.

VBA Expressions 1.2.0 cannot evaluate expressions that contain variables on LibreOffice 24.2: there, an optional `String` parameter with an empty default raises "Data type mismatch" when the caller leaves it out and the procedure compares it. Expressions without variables are not affected. Version 1.2.1 sets those parameters explicitly.

### Python

`ASF.oxt` carries a Python script that runs on LibreOffice's own Python (on Linux, the system's, where numpy and pandas install as usual). It gives what the Python add-in gives Excel (`excel/python`): functions of your own, in the modules of a Python folder (`%APPDATA%\ASF\python` on Windows, `~/.config/ASF/python` elsewhere), marked with `@export` from the `asfpy` module:

```python
# stats.py in the Python folder
from asfpy import export

@export
def scale(values: list, factor: float):
    return [v * factor for v in values]
```

A script gets `py` when the host gives it: level 1 calls exported functions, level 2 imports any module too. Python can't be sandboxed, so give level 2 only to scripts you trust.

```vb
GlobalScope.BasicLibraries.LoadLibrary "ASF"
engine.InjectHostObject "py", LOpython.PyRoot(1)
MsgBox engine.Run(engine.Compile("return py.call('stats.scale', [1, 2, 3], 10).length;"))   ' => 3
r = LOpython.PyCall("stats.scale", Array(1, 2), 10)                ' from Basic: rows from 1
MsgBox r(1, 1)                                                      ' => 10
LOpython.PySetFolder "/home/me/python"                              ' another folder ("" for the default)
MsgBox LOpython.PyReload()                                          ' after the modules changed
```

In a cell, `=PY_CALL("stats.scale"; A1:A3; 10)` (an array formula for a result over several cells) calls the same function once the `ASF` library is loaded, for example by a macro on *Tools > Customize > Events > Start Application*. Calc takes no error value from a Basic function, so an unknown name shows the text `#NAME?` and a Python exception `#VALUE!`; `LOpython.PyLastError()` says why. `PY_CALL` takes up to 15 arguments. The script runs only from the extension, never from a document. `tools/run_python.py` checks `PY_CALL` as Calc calls it, from Basic; in Calc cells it hasn't been checked yet, since a headless LibreOffice runs no Basic function from a cell.

## Differences from the VBA edition

- No Excel object model: `ThisWorkbook`, `Worksheet` and `Range` are not available. COM prototypes target Excel object types, so the Excel prototype libraries in `test/` do not apply.
- `TypeName` of objects, default members, `Debug.Print`, `CallByName` on Collections and UNO objects, and creating classes of another library with `New` behave differently in LibreOffice Basic; the edition handles them internally, apart from the factories above.

## Building

```sh
python3 libreoffice/tools/build.py
```

writes the library to `libreoffice/dist/ASF/` and the extension to `libreoffice/dist/ASF.oxt`; copy that file to `libreoffice/ASF.oxt` to publish a new build.

Code in `src/` must not resize an array in a one-line `If` (`If n > UBound(a) Then ReDim Preserve a(...)`): LibreOffice then drops later writes to that array without an error. Use a block `If`.

- [`tools/build.py`](tools/build.py): the rewrite passes and the packaging.
- [`compat/LOhelper.bas`](compat/LOhelper.bas): helpers the rewritten code calls, the factories and the output functions.
- [`compat/VBAEhelper.bas`](compat/VBAEhelper.bas): creates `VBAExpressionsLib` evaluators.
- [`compat/LOpython.bas`](compat/LOpython.bas), [`compat/ASF_PyObject.bas`](compat/ASF_PyObject.bas) and [`python/asfpy_lo.py`](python/asfpy_lo.py): Python. `build.py` puts `excel/python/xll/boot.py` into the script; `tools/run_python.py` checks it in a temporary profile.
- [`overrides/ASF_Map.bas`](overrides/ASF_Map.bas): hand-written `ASF_Map`; keep it in step with `src/ASF_Map.cls`.

The extension holds ASF 4 only; `build.old_library_modules()` builds the classes of `src/old/` (ASF 3.2.1) for the benchmarks.
- [`extension-page.html`](extension-page.html): the description for the extension's page on extensions.libreoffice.org.

## Tests

```sh
python3 libreoffice/tools/run_tests.py [regex|runner|all|bench|asfbench|versionbench] [--keep-log] [--vba-expressions=VBAExpressions.oxt]
```

builds the library, starts a headless LibreOffice with a throw-away profile and runs `test/RegexTests.bas` and `test/TestRunner.bas`; `bench` runs the regex engine benchmarks of `test/RegexBench.bas` `asfbench` the compiler and VM benchmarks of `test/AsfBench.bas`, and `versionbench` the workloads of `test/VersionBench.bas` on ASF 3.2.1 (`src/old/`, built with the same passes and loaded next to ASF 4) and ASF 4; they print the milliseconds per call. It needs LibreOffice and its Python bindings (`python3-uno`). The VBA expression tests run when the VBA Expressions extension, or its unpacked `VBAExpressionsLib` folder, is given. Tests that need Excel or the UDFs of `src/UDFunctions.cls` are skipped.
