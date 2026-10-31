# ASF for LibreOffice

The LibreOffice edition of ASF: the same engine, language, regex engine and module system, packaged as a LibreOffice Basic library named `ASF` and shipped as the `ASF.oxt` extension.

It is generated from the VBA sources in [`src/`](../src), which stay the single source of truth: [`tools/build.py`](tools/build.py) rewrites each class for LibreOffice Basic (`Option VBASupport 1`). Every rewrite works around one documented difference between VBA and LibreOffice Basic, so a fix or feature in `src/` reaches both editions.

## Requirements

- LibreOffice 7.5 or later (tested on 24.2).
- For `@(...)` VBA expressions only: the [VBA Expressions extension](https://extensions.libreoffice.org/en/extensions/show/70059) 1.2.1 or later (`VBAExpressionsLib`). ASF works without it until a script evaluates an `@(...)` expression.

## Installation

*Tools > Extension Manager > Add...* and pick `ASF.oxt`, or from a terminal:

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

## Differences from the VBA edition

- No Excel object model: `ThisWorkbook`, `Worksheet` and `Range` are not available. COM prototypes target Excel object types, so the Excel prototype libraries in `test/` do not apply.
- `TypeName` of objects, default members, `Debug.Print`, `CallByName` on Collections and UNO objects, and creating classes of another library with `New` behave differently in LibreOffice Basic; the edition handles them internally, apart from the factories above.

## Building

```sh
python3 libreoffice/tools/build.py
```

writes the library to `libreoffice/dist/ASF/` and the extension to `libreoffice/dist/ASF.oxt`.

Code in `src/` must not resize an array in a one-line `If` (`If n > UBound(a) Then ReDim Preserve a(...)`): LibreOffice then drops later writes to that array without an error. Use a block `If`.

- [`tools/build.py`](tools/build.py): the rewrite passes and the packaging.
- [`compat/LOhelper.bas`](compat/LOhelper.bas): helpers the rewritten code calls, the factories and the output functions.
- [`compat/VBAEhelper.bas`](compat/VBAEhelper.bas): creates `VBAExpressionsLib` evaluators.
- [`overrides/ASF_Map.bas`](overrides/ASF_Map.bas): hand-written `ASF_Map`; keep it in step with `src/ASF_Map.cls`.
- [`extension-page.html`](extension-page.html): the description for the extension's page on extensions.libreoffice.org.

## Tests

```sh
python3 libreoffice/tools/run_tests.py [regex|runner|all|bench] [--keep-log] [--vba-expressions=VBAExpressions.oxt]
```

builds the library, starts a headless LibreOffice with a throw-away profile and runs `test/RegexTests.bas` and `test/TestRunner.bas`; `bench` runs the regex engine benchmarks of `test/RegexBench.bas` and prints the milliseconds per call. It needs LibreOffice and its Python bindings (`python3-uno`). The VBA expression tests run when the VBA Expressions extension, or its unpacked `VBAExpressionsLib` folder, is given. Tests that need Excel or the UDFs of `src/UDFunctions.cls` are skipped.
