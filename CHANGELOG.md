# Changelog

All notable changes for ASF. This file combines the release notes from the project's releases.

## [v3.0.1] - 2026-02-04
https://github.com/ECP-Solutions/ASF/releases/tag/v3.0.1

## Summary
ASF v3.0.1 adds a comprehensive `Math` object providing JavaScript-compatible mathematical constants and functions. The Math object exposes 28 functions and 2 constants accessible through property syntax (e.g., `Math.sin(x)`, `Math.PI`), enabling scientific computing and advanced numerical operations within VBA Advanced Scripting.

---

## Highlights

- **Added**
    - Math object — JavaScript-compatible mathematical namespace with constants and methods:
        ```javascript
        // ── Mathematical constants ─────────────────────────
        Math.E;         // Euler's number ≈ 2.718281828
        Math.PI;        // Pi ≈ 3.141592654

        // ── Trigonometric functions ────────────────────────
        Math.sin(Math.PI / 2);      // => 1
        Math.cos(Math.PI);          // => -1
        Math.tan(Math.PI / 4);      // => 1

        Math.asin(1);               // => 1.5708 (π/2)
        Math.acos(0);               // => 1.5708 (π/2)
        Math.atan(1);               // => 0.7854 (π/4)
        Math.atan2(1, 1);           // => 0.7854 (π/4)

        // ── Hyperbolic functions ───────────────────────────
        Math.sinh(1);               // => 1.1752
        Math.cosh(1);               // => 1.5431
        Math.tanh(1);               // => 0.7616

        Math.asinh(1);              // => 0.8814
        Math.acosh(2);              // => 1.3170
        Math.atanh(0.5);            // => 0.5493

        // ── Exponential & logarithmic ──────────────────────
        Math.exp(2);                // => 7.389 (e²)
        Math.expm1(0.1);            // => 0.1052 (e^x - 1)

        Math.log(Math.E);           // => 1 (natural log)
        Math.log10(100);            // => 2 (base-10 log)
        Math.log2(8);               // => 3 (base-2 log)
        Math.log1p(0.1);            // => 0.0953 (ln(1 + x))

        // ── Power & roots ──────────────────────────────────
        Math.pow(2, 8);             // => 256
        Math.sqrt(16);              // => 4
        Math.cbrt(27);              // => 3 (cube root)
        Math.hypot(3, 4);           // => 5 (√(3² + 4²))
        Math.hypot(1, 2, 2);        // => 3 (√(1² + 2² + 2²))

        // ── Rounding ───────────────────────────────────────
        Math.floor(4.7);            // => 4
        Math.ceil(4.1);             // => 5
        Math.round(4.5);            // => 4 (rounds to even)
        Math.round(5.5);            // => 6

        // ── Sign & absolute value ──────────────────────────
        Math.abs(-5);               // => 5
        Math.sign(-10);             // => -1
        Math.sign(0);               // => 0
        Math.sign(42);              // => 1

        // ── Min/max (variadic) ─────────────────────────────
        Math.max(1, 5, 3, 9, 2);    // => 9
        Math.min(1, 5, 3, 9, 2);    // => 1
        ```

    - Practical examples:
        ```javascript
        // ── Distance calculation ───────────────────────────
        fun distance(x1, y1, x2, y2) {
            return Math.hypot(x2 - x1, y2 - y1);
        };
        d = distance(0, 0, 3, 4);   // => 5

        // ── Degree/radian conversion ───────────────────────
        fun toRadians(degrees) {
            return degrees * Math.PI / 180;
        };
        fun toDegrees(radians) {
            return radians * 180 / Math.PI;
        };
        angle = toRadians(90);      // => 1.5708
        degrees = toDegrees(Math.PI); // => 180

        // ── Circle area ────────────────────────────────────
        fun circleArea(radius) {
            return Math.PI * Math.pow(radius, 2);
        };
        area = circleArea(5);       // => 78.5398

        // ── Statistical operations ─────────────────────────
        fun mean(numbers) {
            sum = numbers.reduce(fun(acc, x) { return acc + x; }, 0);
            return sum / numbers.length;
        };
        fun variance(numbers) {
            avg = mean(numbers);
            return mean(numbers.map(fun(x) {
                return Math.pow(x - avg, 2);
            }));
        };
        fun stddev(numbers) {
            return Math.sqrt(variance(numbers));
        };

        data = [2, 4, 4, 4, 5, 5, 7, 9];
        avg = mean(data);           // => 5
        std = stddev(data);         // => 2

        // ── Sigmoid function (ML) ──────────────────────────
        fun sigmoid(x) {
            return 1 / (1 + Math.exp(-x));
        };
        s = sigmoid(0);             // => 0.5

        // ── Clamp value to range ───────────────────────────
        fun clamp(value, min, max) {
            return Math.max(min, Math.min(max, value));
        };
        clamped = clamp(150, 0, 100); // => 100
        ```

- **Internal core changes**:
    - **VM** (`ASF_VM.cls`):
        - Math object implemented as a property-dispatch builtin in `EvalMemberNode`
        - Check for `baseLocal.GetValue("name") = "Math"` after builtin method lookup
        - Math constants (`E`, `PI`) evaluated immediately and returned
        - Math methods route through existing Call evaluation with `propN` matching function name
        - All 28 functions implemented using VBA native math functions (`Atn`, `Exp`, `Log`, `Sqr`, etc.)
        - Variadic support for `hypot`, `max`, `min` using `For Each` loops over `evaluated` collection

- **Technical Details**:
    - **Math Constants**:
        ```
        Math.E   → Exp(1)           ≈ 2.718281828459045
        Math.PI  → 4 * Atn(1)       ≈ 3.141592653589793
        ```

    - **Trigonometric Functions** (argument in radians):
        ```
        Math.sin(x)     → Sine
        Math.cos(x)     → Cosine
        Math.tan(x)     → Tangent
        Math.asin(x)    → Arcsine (returns NaN if |x| > 1)
        Math.acos(x)    → Arccosine (returns NaN if |x| > 1)
        Math.atan(x)    → Arctangent
        Math.atan2(y,x) → Two-argument arctangent (quadrant-aware)
        ```

    - **Hyperbolic Functions**:
        ```
        Math.sinh(x)    → Hyperbolic sine
        Math.cosh(x)    → Hyperbolic cosine
        Math.tanh(x)    → Hyperbolic tangent
        Math.asinh(x)   → Inverse hyperbolic sine
        Math.acosh(x)   → Inverse hyperbolic cosine (returns NaN if x < 1)
        Math.atanh(x)   → Inverse hyperbolic tangent (returns NaN if |x| >= 1)
        ```

    - **Exponential & Logarithmic Functions**:
        ```
        Math.exp(x)     → e^x
        Math.expm1(x)   → e^x - 1 (more accurate for small x)
        Math.log(x)     → Natural logarithm (returns NaN if x <= 0)
        Math.log10(x)   → Base-10 logarithm (returns NaN if x <= 0)
        Math.log2(x)    → Base-2 logarithm (returns NaN if x <= 0)
        Math.log1p(x)   → ln(1 + x) (returns NaN if x <= 0)
        ```

    - **Power & Root Functions**:
        ```
        Math.pow(x, y)  → x^y
        Math.sqrt(x)    → Square root
        Math.cbrt(x)    → Cube root (handles negative values)
        Math.hypot(...) → √(x₁² + x₂² + ... + xₙ²) (Euclidean norm)
        ```

    - **Rounding Functions**:
        ```
        Math.floor(x)   → Largest integer ≤ x
        Math.ceil(x)    → Smallest integer ≥ x
        Math.round(x)   → Rounds to nearest integer (banker's rounding)
        ```

    - **Other Functions**:
        ```
        Math.abs(x)     → Absolute value
        Math.sign(x)    → -1, 0, or 1 depending on sign
        Math.max(...)   → Maximum value (variadic)
        Math.min(...)   → Minimum value (variadic)
        ```

    - **Implementation Notes**:
        - All trigonometric functions expect radians (not degrees)
        - `Math.round()` uses VBA's `Round()` (round-to-even)
        - Invalid inputs (e.g., `Math.asin(2)`) return the string `"NaN"`
        - `Math.atan2(y, x)` follows JavaScript convention (y-coordinate first)
        - `Math.cbrt(x)` correctly handles negative values: `Math.cbrt(-8) => -2`
        - Variadic functions (`hypot`, `max`, `min`) accept any number of arguments

    - **Special Cases**:
        ```
        Math.acos(1)           => 0
        Math.acos(-1)          => π (3.14159...)
        Math.asin(1)           => π/2 (1.5708...)
        Math.asin(-1)          => -π/2 (-1.5708...)
        Math.acosh(1)          => 0
        Math.atan2(0, 0)       => 0
        Math.atan2(0, -1)      => π
        Math.atanh(0)          => 0
        Math.ceil(-4.7)        => -4
        Math.floor(-4.7)       => -5
        ```

---

## Usage Examples

```javascript
// Scientific calculator
fun quadraticFormula(a, b, c) {
    discriminant = Math.pow(b, 2) - 4 * a * c;
    
    if (discriminant < 0) {
        return null;  // No real solutions
    };
    
    sqrtDisc = Math.sqrt(discriminant);
    x1 = (-b + sqrtDisc) / (2 * a);
    x2 = (-b - sqrtDisc) / (2 * a);
    
    return [x1, x2];
};

// Solve x² - 5x + 6 = 0
solutions = quadraticFormula(1, -5, 6);
// => [3, 2]

// Trigonometric calculations
fun polarToCartesian(r, theta) {
    x = r * Math.cos(theta);
    y = r * Math.sin(theta);
    return { x: x, y: y };
};
```

---

**Full Changelog**: https://github.com/ECP-Solutions/ASF/compare/v3.0.0...v3.0.1

## [v3.0.0] - 2026-02-03
https://github.com/ECP-Solutions/ASF/releases/tag/v3.0.0

## Summary
ASF v3.0.0 is a major-version release that introduces a full ECMAScript-style module system (`import`/`export`) and adds working-directory builtins (`cwd`, `scwd`) together with a single-call `Execute` entry point on the `ASF` UI. Modules are cached on first load, circular dependencies are detected at load time, and relative paths are resolved against the current working directory. Many of the new features are still in beta, awaiting community support and feedback, or under continuous development.

---

## Breaking Changes

- **File extension adopted: `.vas`.**  
  `.vas` stands for **V**BA **A**dvanced **S**cripting.  
  `ReadModuleSource` resolves `.vas` automatically; bare module names without an extension are tried first as-is, then with `.vas` appended.  
  Existing `.asf` source files must be renamed to `.vas` before use with the module loader.

---

## Highlights

- **Added**
    - `.vas` file extension — the official source-file extension for all VBA Advanced Scripting files:

    - Module system — `import` / `export` statements with ECMAScript semantics:
		```javascript
				// ── Named exports  (math.vas) ──────────────────────
				fun add(a, b) {
					return a + b;
				};

				fun multiply(a, b) {
					return a * b;
				};

				PI = 3.14159;

				export { add, multiply, PI };

				// ── Named imports  (main_math.vas) ─────────────────
				scwd(wd);
				import { add, multiply, PI } from './math.vas';

				result = add(5, 3);
				area = PI * multiply(5, 5);
				return `5 + 3 = ${result}, Circle area: ${area}`;
				// => "5 + 3 = 8, Circle area: 78.53975"
		```

		```javascript
				// ── Default export  (calculator.vas) ───────────────
				fun Calculator() {
					return {
						add: fun(a, b) { return a + b; },
						subtract: fun(a, b) { return a - b; }
					};
				};

				export default Calculator;

				// ── Default import  (main_calculator.vas) ──────────
				scwd(wd);
				import calc from './calculator.vas';

				calculator = calc();
				return(calculator.add(10, 5));
				// => "15"
		```

		```javascript
				// ── Namespace import  (utils.vas) ──────────────────
				fun formatName(first, last) {
					return first + ' ' + last;
				};

				fun uppercase(str) {
					return str.toUpperCase();
				};

				export { formatName, uppercase };

				// ── Namespace usage  (main_utils.vas) ──────────────
				scwd(wd);
				import * as utils from './utils.vas';

				name = utils.formatName('John', 'Doe');
				return utils.uppercase(name);
				// => "JOHN DOE"
		```

		```javascript
				// ── Mixed: default + named  (lib.vas) ──────────────
				fun helper() {
					return 'Helper function';
				};

				fun main() {
					return 'Main function';
				};

				VERSION = '1.0.0';

				export default main;
				export { helper, VERSION };

				// ── Mixed imports  (app.vas) ────────────────────────
				scwd(wd);
				import mainFunc, { helper, VERSION } from './lib.vas';

				return `${mainFunc()} | ${helper()} | Version: ${VERSION}`;
				// => "Main function | Helper function | Version: 1.0.0"
		```

    - Aliasing with `as` in both import and export specifier lists:
		```javascript
				import { add as sum } from './math.vas';
				export { localName as publicName };
		```

    - Working-directory builtins (`cwd` / `scwd`):
		```javascript
				// scwd(path)  — set current working directory
				// cwd()       — return current working directory
				scwd(wd);                       // set to injected path
				currentPath = cwd();            // read it back
				// currentPath === wd
		```

    - `Execute(filePath)` method on the `ASF` façade — single call to read, compile, and run a `.vas` file and return its result:
		```vba
				Dim eng As New ASF
				eng.InjectVariable "wd", ThisWorkbook.path
				result = eng.Execute(ThisWorkbook.path & "\main_math.vas")
				' result === "5 + 3 = 8, Circle area: 78.53975"
		```

- **Internal core changes**:
    - **Parser** (`ASF_Parser.cls`):
        - Four identifiers are now intercepted during tokenization and emitted as `["KEYWORD", …]` tokens instead of `["IDENT", …]`: `import`, `export`, `default`, `as`
        - Matching is case-sensitive; only the lower-case forms are recognised as keywords

    - **Compiler** (`ASF_Compiler.cls`):
        - New instance properties `CurrentModulePath` and `IsModuleMode` set by the VM loader before each module compilation
        - New `ParseImportStatement` method — parses all five supported import forms and produces an `Import` AST node
        - New `ParseExportStatement` method — parses named-brace, default-expression, and declaration export forms and produces an `Export` AST node
        - `CompileProgram` main loop now checks for `KEYWORD / import` and `KEYWORD / export` tokens before falling through to the existing statement collector; matched statements are added directly to `stmtsAST` via `GoTo Compiler_MainLoop`

    - **VM** (`ASF_VM.cls`):
        - New builtins `cwd` and `scwd` in the Call-case builtin dispatcher; both read/write `GLOBALS_.CURRENT_MODULE_PATH`
        - New `ExecModImport` — resolves module path, calls `LoadModule`, then binds default, namespace, or named imports into the caller's scope; named-export values are resolved through `EvalExprNode` (Variable fallback to `gFuncTable`) so that top-level `fun` declarations export correctly
        - New `ExecModExport` — writes named or default export values into `gModuleExports` for the currently executing module path
        - New `LoadModule` — orchestrates the full module lifecycle: circular-dependency check against `gLoadingModules`, source read, compilation via a fresh `ASF_Compiler` instance, scope creation, statement execution, pending-export post-processing, and default-export extraction; caches the result in `gModuleRegistry`
        - New `ResolveModulePath` — for paths beginning with `./` or `../`, prepends the current working directory (normalised to forward-slash); otherwise returns the path unchanged
        - New `ReadModuleSource` — delegates to `ResolveModulePath`, appends `.vas` if the file does not exist without it, raises error 9012 if still missing, then delegates to `ReadTextFile`

    - **Globals** (`ASF_Globals.cls`):
        - `gModuleRegistry` (`ASF_Map`) — caches loaded module objects keyed by resolved path
        - `gModuleExports` (`ASF_Map`) — maps each module path to its live exports map during execution
        - `gLoadingModules` (`Collection`) — stack of paths currently being loaded; used for circular-dependency detection
        - `CURRENT_MODULE_PATH` (`String`) — the working directory used by `ResolveModulePath` and read/written by `cwd`/`scwd`
        - All four members are initialised in `ASF_InitGlobals`

    - **ASF UI** (`ASF.cls`):
        - New `Execute(filePath As String)` method — reads, compiles, and runs a `.vas` file in a single call; returns the program output
        - New `WorkingDir` property (get/set) — direct access to `GLOBALS_.CURRENT_MODULE_PATH`
        - New `ClearModuleCache` method — resets `gModuleRegistry`, `gModuleExports`, `gLoadingModules`, and `CURRENT_MODULE_PATH`
        - New `ReadTextFile(filePath)` method — exposes the VM's binary-stream file reader

- **Technical Details**:
    - **Import AST Node**:
		```
				Import {
				  type:              "Import"
				  source:            String              // raw string-literal value from 'from' clause
				  defaultImport:     String              // present only for default-import forms
				  namespaceImport:   String              // present only for * as X forms
				  namedImports:      Collection          // present only when { … } specifiers exist;
				}                                        //   each item is an ImportSpecifier
		```

    - **ImportSpecifier AST Node**:
		```
				ImportSpecifier {
				  type:      "ImportSpecifier"
				  imported:  String      // name as it appears in the exporting module
				  local:     String      // name bound in the importing scope (differs when 'as' used)
				}
		```

    - **Export AST Node** (three shapes):
		```
				// default form
				Export {
				  type:        "Export"
				  isDefault:   True
				  expression:  ExprNode     // parsed expression after 'default'
				}

				// named-brace form
				Export {
				  type:            "Export"
				  isDefault:       False
				  namedExports:    Collection     // each item is an ExportSpecifier
				}

				// declaration form  (export fun …)
				Export {
				  type:               "Export"
				  isDefault:          False
				  declarationType:    "function"
				  declarationName:    String      // function name; main loop re-parses the declaration
				}
		```

    - **ExportSpecifier AST Node**:
		```
				ExportSpecifier {
				  type:       "ExportSpecifier"
				  local:      String      // name in the module's own scope
				  exported:   String      // name exposed to importers (differs when 'as' used)
				}
		```

    - **Module Object** (stored in `gModuleRegistry`):
		```
				Module (ASF_Map) {
				  path:            String      // resolved absolute path used as cache key
				  loaded:          Boolean     // True after full execution completes
				  exports:         ASF_Map     // live named-export bindings (name → value)
				  defaultExport:   Variant     // present only when module contains 'export default'
				}
		```

    - **New Keywords** (tokenised as `["KEYWORD", value]`):
		```
				"import"   "export"  "default"   "as"
		```

    - **Error Messages**:
        - `"Unexpected end after import"` — import statement truncated (Compiler, #8001)
        - `"Expected 'as' after * in import"` — namespace import missing `as` (Compiler, #8002)
        - `"Expected identifier after 'as'"` — `as` not followed by a name (Compiler, #8003 / #8004 / #8011)
        - `"Invalid import syntax"` — unrecognised token after `import` (Compiler, #8005)
        - `"Expected 'from' in import statement"` — missing `from` clause (Compiler, #8006)
        - `"Expected string literal for module path"` — non-string after `from` (Compiler, #8007)
        - `"Unexpected end after export"` — export statement truncated (Compiler, #8010)
        - `"Expected function name after 'fun'"` — `export fun` not followed by name (Compiler, #8012)
        - `"Invalid export syntax"` — unrecognised token after `export` (Compiler, #8013)
        - `"Export 'X' not found in module 'Y'"` — named import references missing key (VM, #9001)
        - `"Circular dependency detected: X"` — module re-entered while still loading (VM, #9010)
        - `"Failed to load module: X"` — source read raised an error (VM, #9011)
        - `"Module file not found: X"` — path does not exist after `.vas` fallback (VM, #9012)

---

**Full Changelog**: https://github.com/ECP-Solutions/ASF/compare/v2.0.8...v3.0.0

## [v2.0.8] - 2026-02-01
https://github.com/ECP-Solutions/ASF/releases/tag/v2.0.8

## Summary
ASF v2.0.8 introduces modern JavaScript-style spread/rest operators (`...`) and array destructuring assignments, significantly enhancing array manipulation capabilities and function parameter handling.

---

## Highlights

- **Added**
    - Spread/rest operator (`...`) support:
```javascript
        // Spread in array literals
        arr1 = [1, 2, 3];
        arr2 = [0, ...arr1, 4, 5]; // => [0, 1, 2, 3, 4, 5]
        
        // Spread in function calls
        fun add(a, b, c) { return a + b + c; };
        numbers = [1, 2, 3];
        result = add(...numbers); // => 6
        
        // Rest parameters in functions
        fun sum(first, ...rest) {
            total = first;
            rest.forEach(fun(n) { total = total + n });
            return total;
        };
        sum(1, 2, 3, 4, 5); // => 15
        
        // Spread strings into character arrays
        chars = [...'hello']; // => ['h', 'e', 'l', 'l', 'o']
```
    
    - Array destructuring assignments:
```javascript
        // Basic destructuring
        [a, b, c] = [1, 2, 3]; // a=1, b=2, c=3
        
        // With rest element
        [first, ...rest] = [1, 2, 3, 4, 5];
        // first=1, rest=[2, 3, 4, 5]
        
        // Fewer targets than elements
        [x, y] = [10, 20, 30, 40]; // x=10, y=20
        
        // More targets than elements (assigns Empty)
        [p, q, r] = [100, 200]; // p=100, q=200, r=Empty
        
        // Practical use cases
        [a, b] = [b, a]; // Swap variables
        [head, ...tail] = myArray; // Extract head and tail
        
        // Combine spread and destructuring
        arr1 = [1, 2];
        arr2 = [3, 4];
        [first, ...combined] = [...arr1, ...arr2];
        // first=1, combined=[2, 3, 4]
```

- **Internal core changes**:
    - **Parser** (`ASF_Parser.cls`):
        - Added tokenization for `...` spread/rest operator
        - New token type: `"SPREAD"` with value `"..."`
        
    - **Compiler** (`ASF_Compiler.cls`):
        - Added `IsSpreadToken()` helper function for spread operator detection
        - Rest parameter support in function and method definitions
        - Spread operator handling in array literal compilation (`ParseArrayLiteral`)
        - Spread operator support in function call argument processing
        - Array destructuring pattern detection in `ParseStatementTokensToAST`
        - New AST node type: `"ArrayDestructuring"` with targets collection and optional rest target
        - Compile-time validation:
            - Rest parameter must be last in function signatures
            - Only one rest parameter allowed per function
            - Rest element must be last in destructuring patterns
            - Multiple rest elements not allowed in destructuring patterns
        
    - **VM** (`ASF_VM.cls`):
        - Spread operator evaluation in `EvalArrayNode` with type handling
        - Array expansion for arrays, strings, and scalar values
        - Spread operator support in function call argument expansion
        - Rest parameter handling in function calls with automatic array creation
        - New `"ArrayDestructuring"` case handler in `ExecuteStmtNode`

- **Technical Details**:
    - **Spread/Rest Token Structure**:
```
        Token: ["SPREAD", "..."]
```
    
    - **Rest Parameter AST**:
```
        Function/Method node {
          params: Collection of parameter names
          restParam: String (optional)
          ...
        }
```
    
    - **Array Destructuring AST**:
```
        ArrayDestructuring {
          targets: Collection of Variable nodes
          restTarget: String (optional)
          source: Expression node
        }
```
    
    - **Error Messages**:
        - `"Rest parameter must be last parameter"` - Rest parameter not in final position
        - `"Multiple rest parameters not allowed"` - More than one rest parameter in function
        - `"Rest element must be last in destructuring"` - Rest element not in final position
        - `"Multiple rest elements in destructuring"` - More than one rest element in pattern
        - `"Expected identifier after ... in destructuring"` - Invalid rest syntax
        - `"Cannot destructure non-array value"` - Runtime type validation failure

---

**Full Changelog**: https://github.com/ECP-Solutions/ASF/compare/v2.0.7...v2.0.8

## [v2.0.7] - 2026-01-28
https://github.com/ECP-Solutions/ASF/releases/tag/v2.0.7

## Summary

ASF v2.0.7 is a minor improvement for `ASF` core ergonomic.

---

## Highlights

- **Improved**: 
	- ASF: the `Run` method now returns the result of the compiled program being executed.
		```vb
		Private Sub ForEachTest()
 		   Dim result As Variant
 		   Dim engine As ASF: Set engine = New ASF
 		   
 		   With engine
		        result = .Run(.Compile("o = {x: 10, y: 20}; s = 0; o.forEach(fun(v) { s = s + v }); return(s);"))
		    End With
		    Set engine = Nothing
		End Sub
		```


---
**Full Changelog**: https://github.com/ECP-Solutions/ASF/compare/v2.0.6...v2.0.7

## [v2.0.6] - 2026-01-23
https://github.com/ECP-Solutions/ASF/releases/tag/v2.0.6

## Summary

ASF v2.0.6 represent an improvement for native objects ergonomics.

---

## Highlights

- **Improved**: 
	- VM: users can now call the `foreach` method over an objet.
		```js
		s=0; c=0; 
		foreach({math: 85, english: 92, science: 78}, fun(val, key){ s = s + val; c += 1 }); 
		return('Average: ' + s/c); // ==> Average: 85
		```


---
**Full Changelog**: https://github.com/ECP-Solutions/ASF/compare/v2.0.5...v2.0.6

## [v2.0.5] - 2026-01-23
https://github.com/ECP-Solutions/ASF/releases/tag/v2.0.5

## Summary

ASF v2.0.5 represent an improvement for VMs collaboration through variable injection.

---

## Highlights

- **Improved**: 
	- ASF: users can now inject `ASF_Map` and `ASF_RegexEngine` objects at VM time.


---
**Full Changelog**: https://github.com/ECP-Solutions/ASF/compare/v2.0.4...v2.0.5

## [v2.0.4] - 2026-01-22
https://github.com/ECP-Solutions/ASF/releases/tag/v2.0.4

## Summary

ASF v2.0.4 is a code clean-up for the library.

---

## Highlights

- **Improved**: 
	- Code: Rubberduck code inspection.


---
**Full Changelog**: https://github.com/ECP-Solutions/ASF/compare/v2.0.3...v2.0.4

## [v2.0.3] - 2026-01-21
https://github.com/ECP-Solutions/ASF/releases/tag/v2.0.3

## Summary

ASF v2.0.3 is a spot on update for the parser.

---

## Highlights

- **Improved**: 
	- Parser: the tokenizer now allows (liberally) the use of double quotes for defining literal strings.


---
**Full Changelog**: https://github.com/ECP-Solutions/ASF/compare/v2.0.2...v2.0.3

## [v2.0.2] - 2026-01-19
https://github.com/ECP-Solutions/ASF/releases/tag/v2.0.2

## Summary

ASF v2.0.2 is a hot fix for the VM.

---

## Highlights

- **Fixed**: 
	- VM: fixed LValue resolution for mutating array methods.
	- VM: fixed QuickSort can not sort array of objects.


---
**Full Changelog**: https://github.com/ECP-Solutions/ASF/compare/v2.0.1...v2.0.2

## [v2.0.1] - 2026-01-17
https://github.com/ECP-Solutions/ASF/releases/tag/v2.0.1

## Summary

ASF v2.0.1 is an improvement to the VM when dealing with classes.

---

## Highlights

- **Fixed**: 
	- VM: fixed out of stack space error when dealing with classes and polymorphism. So, now is safe to execute code like this:
		```js
		class Printer {
			print(doc) { return 'Printing: ' + doc; };
		};
		class ColorPrinter extends Printer {
			print(doc) { return 'Color printing: ' + doc; };
		};
		class LaserPrinter extends Printer {
			print(doc) { return 'Laser printing: ' + doc; };
		};
		fun printDocument(printer, doc) {
			return printer.print(doc);
		};
		p1 = new Printer();
		p2 = new ColorPrinter();
		p3 = new LaserPrinter();
		result = [printDocument(p1, 'Doc1'), printDocument(p2, 'Doc2'), printDocument(p3, 'Doc3')].join(' | ');
		return result; // => Printing: Doc1 | Color printing: Doc2 | Laser printing: Doc3
		```


---
**Full Changelog**: https://github.com/ECP-Solutions/ASF/compare/v2.0.0...v2.0.1

## [v2.0.0] - 2026-01-17
https://github.com/ECP-Solutions/ASF/releases/tag/v2.0.0

## Summary

ASF v2.0.0 is a significant improvement over the previous version. This version adds support for classes, object literal methods, debug tracing, and extends the coverage of the test suite. Also the documentation is now more complete.

---

## Highlights

-  **Added** 
    - `let` keyword support
    - Bitwise or (`|`) operator
    - `undefined` data type
    - Support for classes:
		```js
		class Shape {
		  field x = 0, y = 0, color = 'black';
		  constructor(x, y) {
			this.x = x;
			this.y = y;
		  };
		  getPosition() {
			return this.x + ',' + this.y;
		  };
		};
		class Circle extends Shape {
		  field radius = 1;
		  constructor(x, y, r) {
			super(x, y);
			this.radius = r;
			this.color = 'red';
		  };
		  getArea() {
			return 3.14159 * this.radius * this.radius;
		  };
		};
		let circle = new Circle(10, 20, 5);
		print('Position: ' + circle.getPosition()); // => 'Position: 10,20'
		print('Color: ' + circle.color); // => 'Color: red'
		print('Area: ' + circle.getArea()); // => 'Area: 78.53975'
		```
    - Objects methods:
	```js
	o = {apple: 10, banana: 20, cherry: 30}; result = o.filter(fun(val, key) { return key.startsWith('a') }); 
	return result; // => { apple: 10 }
	```
    - The `ASF.ReadTextFile` method allows users to get source code from text files.
    - Language reference documentation.

- **Internal core change**:
	- **VM**:
		- Improved `step` in `for` node: now supports compound assignments
		- Users can now inspect the call stack through debug tracing.
		
			```vb
			Dim ASF_ As New ASF
			Dim code As String
			
			' Enable call tracing
			ASF_.EnableCallTrace = True
			
			code = "fun add(a, b) { return a + b; };" & vbCrLf & _
				   "fun multiply(a, b) { return a * b; };" & vbCrLf & _
				   "x = add(3, 4);" & vbCrLf & _
				   "y = multiply(x, 3);" & vbCrLf & _
				   "print(y)"
					
			Dim idx As Long
			idx = ASF_.Compile(code)
			ASF_.Run idx
			
			' Print the call stack trace
			Debug.Print "=== Call Stack Trace ==="
			Debug.Print ASF_.GetCallStackTrace()
			
			' Clear for next run
			ASF_.ClearCallStack
			```
			The above code will print
			```
			=== Call Stack Trace ===
			CALL: add(3, 4) -> 7
			CALL: multiply(7, 3) -> 21
			```
- **Fixed**: 
	- Compiler: fixed an error that prevented `Collection` variables from initializing correctly, causing a fault at compilation phase.
	- Compiler: fixed nested try-catch parsing issue.
	- Compiler: fixed `try` statement requiring a mandatory `catch` block.
	- Map container: fixed nested maps cloning bug.
	- Parser: fixed compound assignment bug.
	- VM: fixed use of non-existent property when accessing `ASF_ScopeStack` keys.
	- VM: fixed bug in `typeOf` returning numeric types when invoked with `null` values.
	- VM: fixed switch statement executing all cases even when a match was found.
	- Regex: fixed line of statements compiled as comment in VBA.
	- Regex: fixed greedy sub quantifier.

---
**Full Changelog**: https://github.com/ECP-Solutions/ASF/compare/v1.0.7...v2.0.0

## [v1.0.7] - 2025-12-30
https://github.com/ECP-Solutions/ASF/releases/tag/v1.0.7

## Summary

This ASF v1.0.7 release is focused in regex engine and ergonomics improvements. 

---

## Highlights

-  **Added** 
    - Native regex engine support for:
		- Inline flags (`(?i:...)`, `(?-i:...)`, `(?s:...)`, `(?-s:...)`): only group scoped flags.

- **Internal core change**:
	- Parser & compiler:
		- The `regex` constructor now accepts arguments to return an initialized instance:`re=regex(<args>)`
		- Fix named capture handling: named captures are now stored as arrays with their name and value, and the replacement logic retrieves the value correctly from the array.
	- **VM**:
		- Variable injection support: allow dynamic injection of variables into the program scope before execution. This enables external code to set variables that will be available during program runtime. VBA `Collections` are converted to ASF internal array representation. For other type of objects, ASF only store its type.
		
			```vb
			Dim engine As ASF
			Dim pidx As Long
			Dim coll As New Collection
			
			Set engine = New ASF
			coll.Add "We ": coll.Add "can do ": coll.Add "more with ": coll.Add "ASF!"
			With engine
				.InjectVariable "a", coll
				pidx = .Compile("txt = ''; a.forEach(fun(x){ txt += x }); return(txt);")
				.Run pidx
				actual = CStr(.OUTPUT_) '--> We can do more with ASF!
			End With
			```
- **Fixed**: 
	- Regex: fixed nested groups error.

---
**Full Changelog**: https://github.com/ECP-Solutions/ASF/compare/v1.0.6...v1.0.7

## [v1.0.6] - 2025-12-27
https://github.com/ECP-Solutions/ASF/releases/tag/v1.0.6

## Summary

ASF v1.0.6 now features a powerful native regular expression engine and a large set of JavaScript-style string methods — bringing JavaScript-like text processing and template literals natively to ASF scripts.

---

## Highlights

-  **Added** 
    - Native regex engine with support for:
		- Anchors, character classes and ranges, escapes (`\d`, `\D`, `\w`, `\s`, etc.).
		- Greedy / lazy / **possessive** quantifiers and **atomic groups**.
		- Lookaheads (`(?=...)`, `(?!...)`) and **fixed-width lookbehinds** (`(?<=...)`, `(?<!...)`) with compile-time rejection of variable-width lookbehind expressions.
		- Nested groups, captures and safe pretty-printing of capture results.
		- `RegExp.escape()` functionality (escape an arbitrary string to be used as a literal pattern).
    - **JS-like String API** methods implemented as ASF builtins:  
		- `replace`/`replaceAll` support both string and function replacement, and accept `/pattern/flags` slash-regex strings (flags: `g`, `i`, `m`, `s`). Function replacements receive `(match, p1...pn, offset, originalString)` — offset is zero-based.
		```js
		fun replacer(match, p1, p2, p3, offset, string){return [p1,p2,p3].join(' - ');} 
		newString = 'abc12345#$*%'.replace('/(\\D*)(\\d*)(\\W*)/g', replacer);
		```
- **Internal core change**:
	- Parser & compiler:
		- String templates (backtick syntax) with `${...}` placeholders:
			- Full expression support inside `${...}`; nested `${...}` supported.
			- Literal parts preserve escaped characters (outside placeholders): `\` escapes "\`", "/", "\".
			- Literal parts preserve escaped characters (inside placeholders): `\` **also** escapes `$`, `{`, `}`.
			- Examples:
			```js
			a='Happy! '; return(`I feel ${a.repeat(3)}`);
			// -> Outputs 'I feel Happy! Happy! Happy! '
			```
		- Template tokenizer replaced with a robust `ParseTemplatePartsFromString` that emits a `Collection` of `Array("LITERAL", text)` and `Array("EXPR", text)` parts — each `${...}` is captured as an `EXPR` exactly as typed respecting escaping rules inside.
		- Recursive-descent expression parser improved so parenthesized expressions consume and allow postfix chaining (member calls, indexing, function calls) after the closing `)`.
- **Fixed**: 
	- Template tokenizer: fixed duplicated literal pieces, correct handling of nested ${...} and escaped characters inside literal parts.

---
**Full Changelog**: https://github.com/ECP-Solutions/ASF/compare/v1.0.5...v1.0.6

## [v1.0.5] - 2025-12-21
https://github.com/ECP-Solutions/ASF/releases/tag/v1.0.5

## Summary

ASF v1.0.5 supports a large number of string functions. The `String` data type is treated as an object from which methods can be invoked.

---

## Highlights

-  **Added** 
    - javascript functions: `length`, `charAt`, `charCodeAt`, `concat`, `endsWith`, `fromCharCode`, `includes`, `indexOf`, `lastIndexOf`, `localeCompare`, `padEnd`, `padStart`, `repeat`, `replace`, `slice`, `split`, `startsWith`, `substring`, `toLowercase`, `toUppercase`, `trim`, `trimStart`, `trimEnd`.  
	
		So users can now write code like this one for advanced string operations
        ```js
			welcome=fun(string){return string.concat('!')}; return('Hello world'.replace('world', welcome('VBA')));
			//-> Outputs 'Hello VBA!'
        ```
    - String templates are now supported.  
         ```js
			a='Happy! '; return(`I feel ${a.repeat(3)}`);
			//-> Outputs 'I feel Happy! Happy! Happy! '
        ```
- **Internal core change**: the builtins methods are now packed by object type.

---
**Full Changelog**: https://github.com/ECP-Solutions/ASF/compare/v1.0.4...v1.0.5

## [v1.0.4] - 2025-12-20
https://github.com/ECP-Solutions/ASF/releases/tag/v1.0.4

## Summary

ASF v1.0.4 has now a syntax more similar to javascript. The Framework can executed `for-in` and `for-of` methods.

---

## Highlights

-  **Added** 
    - javascript `typeof` operator 
        ```js
        f = fun(a) { a = a + 1; return a }; print(typeof f); print(typeof f(1)); //--→PRINT:'function', PRINT:'number'
        ```
    - `for`-`in/of` loop support. 
         ```js
        o = { a:1, b:2 }; keys=[]; for (k in o) { keys.push(k) }; print(keys); //--→PRINT:[ 'a', 'b' ]
        ```
- **Fixed**: extra element appended by the push operation when operating empty arrays.

---
**Full Changelog**: https://github.com/ECP-Solutions/ASF/compare/v1.0.3...v1.0.4


## [v1.0.3] - 2025-12-14
https://github.com/ECP-Solutions/ASF/releases/tag/v1.0.3

## Summary

ASF v1.0.3 can now handle deeply nested indexes. So the lack of deeply index access is now fulfilled.

---

## Highlights

This input is now accepted

```js
a=[1,[[2,3],4],5]; a[2][1][2]=10; print(a);
```

The above will print:

```
PRINT:[ 1, [ [ 2, 10 ], 4 ], 5 ]
```

---
**Full Changelog**: https://github.com/ECP-Solutions/ASF/compare/v1.0.2...v1.0.3

## [v1.0.2] - 2025-12-14
https://github.com/ECP-Solutions/ASF/releases/tag/v1.0.2

## Summary

ASF v1.0.2 can now handle nested indexes. So the bug related to index access are now fixed.

---

## Highlights

This input is now accepted

```js
a=@({{7;8;9}}); a[1][2]=42; print(@(HostCheckSecondNestedElement(a)));
```

Being `HostCheckSecondNestedElement` a function defined in the `UDFunctions.cls` module as

```vb
Public Function HostCheckSecondNestedElement(arr As Variant) As String
    Dim tmpArr As Variant
    Dim expHelper As VBAexpressions
    Set expHelper = New VBAexpressions
    tmpArr = expHelper.ArrayFromString(CStr(arr(0)))
    HostCheckSecondNestedElement = tmpArr(LBound(tmpArr) + 1)
End Function
```

The above will print:

```
'42'
```

Keep in mind that VBA Expressions treat data as `String`.

---
**Full Changelog**: https://github.com/ECP-Solutions/ASF/compare/v1.0.1...v1.0.2

## [v1.0.1] - 2025-12-13
https://github.com/ECP-Solutions/ASF/releases/tag/v1.0.1

## Summary

ASF v1.0.1 can now mutate arrays returned VBA Expressions. This completes the integration phase at the top level. 

---

## Highlights

Users can now inject array expressions and mutates  like this :

```js
a=@({{1;2;3};{4;(5+4);'value'}}); a[1]=2*5; print(a.push(11)); print(a);
```

The console log will show:

```
=== Runtime Log ===
RUN Program: @anon
PRINT:3
PRINT:[ 10, [ 4, 9, 'value' ], 11 ]
```
---

## [v1.0.0] - 2025-12-12
https://github.com/ECP-Solutions/ASF/releases/tag/v1.0.0

## Summary

This is the first stable release of the **Advanced Scripting Framework (ASF)**. ASF v1.0.0 delivers a mature AST-based scripting layer embedded inside VBA with modern language features, shared-write closures, an extensive set of array helpers exposed as bound properties, improved parser robustness, and major performance upgrades. The release is validated by an 85+ test Rubberduck suite.

---

## Highlights

- Full modern array-method suite exposed as **bound properties** (callable as `arr.map`, `arr.filter`, etc.), supporting method chaining and first-class method values.
- Expression-level anonymous functions and top-level functions (`fun(...) { ... }`) with **shared-write closures** (closures capture runtime scopes by reference).
- Parser refactor and robustness: reliable postfix chaining, improved token handling, and `.length` compiled to an internal builtin for consistent semantics.
- Performance improvements: replaced many `Collection`-heavy code paths with fast Variant-array implementations; `sort` now uses in-place QuickSort with comparator support.
- Comprehensive Rubberduck test-suite (85+ tests) covering language features, array helpers, nested arrays, closures, VBExpr integration, and edge cases.

---

## Notable features

- **Array & collection helpers** (complete list):
  - `map`, `filter`, `reduce`, `forEach`, `slice`, `push`, `pop`, `shift`, `unshift`, `concat`, `unique` (deep-aware), `flatten` (depth-aware), `clone`, `toString`/`join`, `delete`, `splice`, `toSpliced`, `at`, `copyWithin`, `entries`, `every`, `find`, `findIndex`, `findLast`, `findLastIndex`, `from`, `includes`, `indexOf`, `lastIndexOf`, `of`, `reverse`, `some`, `sort`, `toReversed`, `toSorted`, `with`.
- **Method-as-property design**: methods are first-class bound objects; `m = a.map` preserves binding and can be invoked later.
- **VBExpr integration**: `@(...)` blocks pass raw VBAexpressions to the VBA-expressions bridge at runtime.
- **Null literal**: `null` supported as a language literal to represent absence of value.
- **Immutable variants**: `toSpliced`, `toReversed`, `toSorted`, `with` — non-mutating convenience APIs.

---

## Bug fixes

- Fixed a parser bug that truncated `if`/block bodies due to premature token consumption — `ParseIfAST` now reliably receives full body tokens.
- Resolved `EvalMemberNode` error when accessing array-returned bases; member handling now returns bound builtin method objects for array methods.
- Fixed nested-array `map` OOB errors and ensured nested arrays are handled recursively and safely.

---

## Performance & implementation notes

- Heavy use of Variant arrays instead of `Collection` where performance matters (array conversion helpers retained for compatibility).
- `sort` / `toSorted` implemented with in-place QuickSort (Hoare partition), supporting optional comparator closures.
- Deep-equality (`DeepEqual`) implemented for structural comparisons used by `unique`, `includes`, and other searches. 

---

## Tests & validation

- Canonical test runner: `test/TestRunner.bas` (85+ Rubberduck tests covering expressions, control flow, closures, array/object literals, method chaining, builtins, VBExpr). Running the test suite is the recommended validation step after importing modules.
- All tests passing at release time.

---

## Breaking changes & migration notes

- **Equality semantics**: Structural (deep) equality is used by default for many helpers (`includes`, `unique`, etc.). Code expecting reference-equality should be updated to use explicit checks.
- **`__option_base`**: Index base (0 or 1) is consistently honored across all array helpers; ensure your runtime `__option_base` is configured if you rely on a specific indexing convention. However, base 1 is the only recommended and tested option. 

---

## Files changed (primary)

- `Compiler.cls` — parser fixes, `ParsePrimaryNode` refactor, postfix chaining extraction (`DoPostfixChaining`), collapsed identifier expansion
- `VM.cls` — runtime method dispatch, builtins, `EvalMemberNode`, QuickSort helpers, LValue writebacks (`ResolveProp` / `WriteToContainer`)
- `ScopeStack.cls`, `Map.cls` — scope/closure helpers and Map semantics
- `TestRunner.bas` — extended Rubberduck suite (canonical tests validating features)

---

## Contributors

- **Lead developer & maintainer:** @ws-garcia 
- Special thanks to @sancarn,  `stdLamda` author, for his extensive and inspirational work on the VBA framework niche. Also thanks to @senipah, for his incredible inspirational work on providing new features to our loved VBA language. 

---

## Upgrade & installation guide

1. Checkout the `ASF-v1.0.0` tag/branch from the project repository.
2. Import modules into your VBA project in the recommended order:
   - `ASF_Map.cls`, `ASF_ScopeStack.cls`, `ASF_Compiler.cls`, `ASF_VM.cls`, `ASF_Globals.cls`, `ASF.cls`, `UDFunctions.cls`,  `VBAexpressions.cls`, `VBAexpressionsScope.cls`
3. Run the test suit by open the [`ASF v0.0.1.xlsm`](/test/ASF v0.0.1.xlsm) or importing the `TestRunner.bas` under Rubberduck to validate the environment and confirm passing tests.
4. Update any user scripts that relied on loose semicolon usage or reference-equality assumptions.

---

## Future work

- Optional strict/reference equality operator (`===`) and configuration for equality semantics.
- Hash-based acceleration for `unique`/`includes` on large collections.
- Expanded documentation: user guide, migration guide, and a performance tuning section.
