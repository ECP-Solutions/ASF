# ASF — Language Syntax Reference

This document summarizes the ASF scripting language supported by the `ASF_Compiler`/`ASF_VM` class modules.

## Lexical
- Identifiers: letters, digits, `_`; case-sensitive at runtime but parsed case-insensitively in some builtins.
- Strings: single quotes: `'hello'`
- Numbers: integer and floating-point (VBA-compatible numeric formats).
- Special token: VBExpression block `@(...)` — evaluates raw VBAexpressions at runtime.

## Statements & Terminators
- Statements are separated by `;` (semicolon). Some semicolons are **required** between statements to remove ambiguity.
- Blocks are delimited by `{ ... }`.

## Expressions
- Binary operators (in precedence order, highest → lowest):
  - `^` (power) — right-associative
  - `*`, `/`, `%`
  - `+`, `-`
  - `<<`, `>>` (shifts)
  - Relational: `<`, `>`, `<=`, `>=`, `==`, `!=`
  - Logical: `&&`, `||`
  - Ternary: `cond ? a : b`
- Unary: `-x`, `!x` (logical NOT)
- Short-circuit semantics: `&&`, `||` short-circuit evaluation.

## Primary forms
- Literals: numbers, strings, `true`, `false`, `null`
- Arrays: `[expr, expr, ...]` — supports nested arrays; array literals create indexable arrays with the interpreter's option base.
- Objects: `{ key: value, ... }` — maps with string keys.
- Identifiers: variables or collapsed forms like `o.a[3].b`
- Calls: `fn(arg1, arg2,..., argn)` or method style `obj.method(args)`
- Anonymous functions: `fun(p1, p2) { ... }` (expression or statement-level)

## Member / Index / Call chaining
- Any primary may be postfixed with `.prop`, `[index]`, or `(args)`. The parser supports chaining like `a.filter(...).reduce(...).slice(...)`.
- Special-case `.length` is compiled to a builtin call `.__len__(x)` at compile-time.

## Builtin functions & methods
(Short list — full API in API.md)
- Array methods (method-style): `.map(fn)`, `.filter(fn)`, `.reduce(fn, initial?)`, `.slice(start, end?)`, `.push(val)`, `.pop()`
- Named builtins: `range(start?, end?, step?)`, `flatten(arr, depth?)`, `clone(value)`, `IsArray(x)`, `IsNumeric(x)`.

## VBExpression integration
- Use `@({...})` to embed a raw VBA expression or block. The parser emits a `VBAexpr` node and the VM uses the integrated VBAExpressions evaluator at runtime.

## Semantics highlights
- **Closures**: capture the lexical `ScopeStack` by reference (shared-write semantics), so nested functions share and mutate outer variables as in JavaScript/Python.
- **Arrays**: 1-based or 0-based behavior follows the interpreter `__option_base` (the VM respects the option base when creating and indexing arrays).
- **Assignment**: supports compound assignments like `+=` and `>>=`; assignment normalization pass expands them into canonical forms in the compiler.

## Error handling
- Parser errors raise `Compiler.Parse...` VB errors with helpful messages.
- The VM catches runtime errors in `try/catch` constructs in ASF script code.

## Examples
- Anonymous function literal:
```js
fun add(a, b) { return a + b }
```

Method chain:
```js
a.filter(fun(x){ return x > 2 }).map(fun(x){ return x*2 }).slice(1,3)
```

For a full API and details on builtins, see API.md.
