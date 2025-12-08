# Advanced Scripting Framework (ASF)
## ![ASF](/docs/assets/img/ASF%20logo.png)
[![Tests (Rubberduck)](https://img.shields.io/badge/tests-Rubberduck-brightgreen)](https://rubberduckvba.com/)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

> Modern scripting power inside classic VBA. Fast to adopt — impossible to ignore. 
> **ASF** gives you useful scripting features (closures, objects, arrays, first-class functions) inside VBA — without leaving the Office ecosystem.

ASF is an embeddable scripting engine written in plain VBA that brings modern language features — first-class functions, anonymous closures, array & object literals, and safe interop with your existing VBA code — to legacy Office apps.

This project provides a production-proven compiler and VM plus a complete test-suite validating semantics and runtime behavior.

---

## Why ASF?
- **Unmatched expressiveness:** Implement complex logic with concise scripts and enrich them with heavyweight VBA code.
- **Safe interoperability:** Delegate numeric and domain-specific work to your existing VBA functions via `@(...)`, we already have [VBA-expressions](https://github.com/ECP-Solutions/VBA-Expressions) embedded!
- **Non-invasive adoption:** Import a few class modules and you’re ready — no COM servers, no external dependencies.
- **Ship scripting to end-users without new runtimes.** Embed scripts into Excel/Access/Word projects and run code dynamically.
- **Readable, debuggable AST-first design.** The Compiler emits Map-based ASTs (human-inspectable). The VM executes those ASTs directly so you can step through behavior and trace problems — no opaque bytecode black box.
- **Enterprise-ready:** Canonical source modules (`ASF_Compiler.cls`, `ASF_VM.cls`) and a full [Rubberduck](https://github.com/rubberduck-vba/Rubberduck) test-suite enable confident audits and CI.
- **Closure semantics you actually expect.** Shared-write closure capture (like JavaScript/Python) keeps behavior intuitive.
- **Progressive optimization path.** Start with a rock-solid AST runtime for correctness — later switch on the performant compact-bytecode fast-path with minimal changes.
- **Designed for real engineering work.** Robust array/object handling, VB-expression passthrough (`@(...)`), and a small host-wrapper for easy integration.

---

## Highlights / Features

- Full expression language: arithmetic, boolean, ternary, short-circuit logic.
- Arrays, objects (Map-like), member access and indexing.
- First-class functions + anonymous functions + closures.
- Control flow: `if` / `elseif` / `else`, `for`, `while`, `switch`, `try/catch`, `break` / `continue`.
- `print(...)` convenience for quick debugging.
- VBA expressions passthrough (`@(...)`) to call into native user defined functions where needed.
- Traceable runtime log via `GLOBALS_.gRuntimeLog` for deep debugging.
- Compact wrapper (`ASF` class) — `Compile` + `Run` are one-liners from host code.

---

## Quick Start
1. Import canonical modules into your VBA project (recommended list below).
2. Optionally initialize globals to register UDFs and share evaluators.
3. Compile and run scripts from your host code.

**Recommended module list:** `ASF.cls`, `ASF_Compiler.cls`, `ASF_VM.cls`, `ASF_Globals.cls`, `ASF_ScopeStack.cls`, `ASF_Parser.cls`, `ASF_Map.cls`, `UDFunctions.cls`, `VBAcallBack.cls`, `VBAexpressions.cls`, `VBAexpressionsScope.cls`.

**Minimal example**

```vb
Dim engine As ASF
Set engine = New ASF
Dim idx As Long
idx = engine.Compile("a = 1; f = fun() { a = a + 1; return a }; print(f()); print(a);")
engine.Run idx
' Inspect engine.GetGlobals.gRuntimeLog
```
---

## Features & Capabilities

- Full AST-based compiler and VM implemented in VBA.
- Function literals (anonymous), named top-level functions, recursion.
- Arrays and objects with literal syntax and `.length` helpers.
- Member access, nested indexing, and LValue semantics for assignments.
- Short-circuit logical operators, ternary operator, compound assignments.
- VB-expression embedding: reuse your VBA libraries seamlessly.
- Pretty-printing, runtime logging and cycle-safe map/collection serialization.

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
