# Advanced Scripting Framework (ASF)
## ![ASF](/docs/assets/img/ASF%20logo.png)
[![Tests (Rubberduck)](https://img.shields.io/badge/tests-Rubberduck-brightgreen)](https://rubberduckvba.com/)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

> Modern scripting power inside classic VBA. Fast to adopt — impossible to ignore. 
> Turn VBA into a full-featured script host. **ASF brings modern scripting ergonomics (first-class functions, closures, arrays & objects, method chaining, builtin helpers like `map`/`filter`/`reduce`, and even VBExpressions) to classic VBA projects — without leaving the Office ecosystem.

ASF is an embeddable scripting engine written in plain VBA that brings modern language features — first-class functions, anonymous closures, array & object literals, and safe interop with your existing VBA code — to legacy Office apps.

This project provides a production-proven compiler and VM plus a complete test-suite validating semantics and runtime behavior.

---

## Why ASF?
- **Seamless bridge** between VBA codebases and modern scripting paradigms.
- **No external runtime** — runs on top of VBA using a compact AST interpreter.
- **Powerful features** not found in any other VBA tool: shared-write closures, expression-level anonymous functions, nested arrays & objects, array helpers, VBExpressions integration, method chaining, and more.
- **Tested** — the comprehensive [Rubberduck](https://github.com/rubberduck-vba/Rubberduck)  test-suite passes across arithmetic, flow control, functions, closures, array/object manipulation and builtin methods.
- **Unmatched expressiveness:** Implement complex logic with concise scripts and enrich them with heavyweight VBA code.
- **Safe interoperability:** Delegate numeric and domain-specific work to your existing VBA functions via `@(...)`, we already have [VBA-expressions](https://github.com/ECP-Solutions/VBA-Expressions) embedded!
- **Readable, debuggable AST-first design.** The Compiler emits Map-based ASTs (human-inspectable). The VM executes those ASTs directly so you can step through behavior and trace problems — no opaque bytecode black box.
- **Closure semantics you actually expect.** Shared-write closure capture (like JavaScript/Python) keeps behavior intuitive.
- **Designed for real engineering work.** Robust array/object handling, VB-expression passthrough (`@(...)`), and a small host-wrapper for easy integration.

---

## Highlights / Features

- Full expression language: arithmetic, boolean, ternary, short-circuit logic.
- Arrays, objects (Map-like), member access and indexing.
- First-class functions + anonymous functions + closures.
- Control flow: `if` / `elseif` / `else`, `for`, `while`, `switch`, `try/catch`, `break` / `continue`.
- Map / Filter / Reduce / Slice / Push / Pop as array methods
- Builtin helpers: `range`, `flatten`, `clone`, `IsArray`, `IsNumeric`, etc.
- Method chaining: `a.filter(...).reduce(...)`
- `print(...)` convenience for quick debugging.
- Pretty-printing for arrays/objects with cycle-safety
- VBA expressions passthrough (`@(...)`) to call into native user defined functions where needed.
- Compact wrapper (`ASF` class) — `Compile` + `Run` are one-liners from host code.
  
---

## Quick Start
1. Import canonical modules into your VBA project (recommended list below).
2. Optionally initialize globals to register UDFs and share evaluators.
3. Compile and run scripts from your host code.

**Recommended module list:** `ASF.cls`, `ASF_Compiler.cls`, `ASF_VM.cls`, `ASF_Globals.cls`, `ASF_ScopeStack.cls`, `ASF_Parser.cls`, `ASF_Map.cls`, `UDFunctions.cls`, `VBAcallBack.cls`, `VBAexpressions.cls`, `VBAexpressionsScope.cls`.

**Examples**

Map & nested arrays:
```vb
Dim engine As ASF
Set engine = New ASF
Dim idx As Long
idx = engine.Compile("a = [1,'x',[2,'y',[3]]];" & _
"b = a.map(fun(x){" & _
  "if (IsArray(x)) { return x }" & _
  "elseif (IsNumeric(x)) { return x*3 }" & _
  "else { return x }" & _
"});" & _
"print(b);")
engine.Run idx '// => [ 3, 'x', [ 6, 'y', [ 9 ] ] ]
```
Chained helpers:
```vb
"a=[1,2,3,4,5]; return(a.filter(fun(x){ return x > 2 }).reduce(fun(acc,x){ return acc + x }, 0));" '// => 12
```
---

## Features & Capabilities

- Full AST-based compiler and VM implemented in VBA.
- Function literals (anonymous), named top-level functions, recursion.
- Arrays and objects with literal syntax and `.length` helpers.
- Member access, nested indexing, and LValue semantics for assignments.
- Short-circuit logical operators, ternary operator, compound assignments.
- VB-expression embedding: reuse your VBA libraries seamlessly.
  
---

## Examples & Patterns

Explore `examples/` (suggested) with scripts converting rules, workflows, or automation into ASF scripts. The test-suite provides dozens of ready-to-run scenarios.

---

## Running the Test Suite

1. Import `tests/TestRunner.bas` Rubberduck test module, or open the `ASF v0.0.1.xlsm` workbook.
2. Ensure [`Rubberduck`](https://rubberduckvba.com/) add-in is available.
3. Run the test module — all canonical tests should pass.

---

## Contributing & Roadmap

- Report bugs or propose features via Issues.
- PRs must include tests covering behavior changes.
- Roadmap: improved diagnostics, optional sandboxing primitives, richer standard library for arrays/strings.

---

## License

MIT — see `LICENSE`.

---

For enterprise or integration help, reach out with a short description of your environment and goals — ASF is intentionally lightweight so it adapts quickly to complex legacy codebases.
