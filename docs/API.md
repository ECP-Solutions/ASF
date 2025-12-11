# ASF — Runtime API & Builtins

This document lists the runtime builtins and VM integration points.

## Embedding & Host API (VBA)
- `ASF` class: primary host object (expose via `Set obj = New ASF`)
  - `Compile(source As String) As Long` — compile source and return program index
  - `Run(programIndex As Long)` — execute compiled program
  - `OUTPUT_` — public property storing last returned value (stringified or raw depending on call)

- `ASF_Globals` — runtime-global container, used to configure array option base, expression evaluator, and logging.
  - `.ASF_InitGlobals` — init defaults
  - `.gExprEvaluator` — VBAExpressions integration (declare UDFs here)
  - `.gRuntimeLog` — internal runtime log used by the tests for `print()` messages

## AST & VM relevant shapes (for contributors)
- AST nodes are implemented as `ASF_Map` objects with a `"type"` key and node-specific keys:
  - `Literal`: keys: `type="Literal"`, `value`
  - `Variable`: `type="Variable"`, `name`
  - `Member`: `type="Member"`, `base` (AST), `prop`
  - `Index`: `type="Index"`, `base` (AST), `index` (AST)
  - `Call`: `type="Call"`, `callee` (AST), `args` (Collection of ASTs)
  - `FuncLiteral`: `type="FuncLiteral"`, `params` (Collection), `body` (Collection of statement ASTs)
  - `VBAexpr`: `type="VBAexpr"`, `expr` (string)
  - `Object`: `type="Object"`, `items` (Collection of pairs)

## Key builtins (method & named)
### Array methods (method-style)
- `map(fn)`  
  Returns a new array with `fn` applied to each element. `fn` receives `(value, index, array)` — index uses interpreter base semantics. Supports nested arrays and returns deep nested structures unchanged unless `fn` mutates.
- `filter(fn)`  
  Returns a new array with elements where `fn` returns truthy.
- `reduce(fn, initial?)`  
  Left-to-right reduction. If `initial` omitted, first element is used as initial accumulator.
- `slice(start, end?)`  
  Non-mutating slice. **Important:** indices respect `__option_base`. `end` is exclusive.
- `push(val...)`  
  Appends one or more values in-place and returns new length.
- `pop()`  
  Removes last element and returns it.

### Named builtins
- `range(stop)` / `range(start, stop)` / `range(start, stop, step)`  
  Returns an array of numeric sequence.
- `flatten(arr)` / `flatten(arr, depth)`  
  Flattens the array. No depth => full flatten. `depth = 0` => no flatten; `depth = 1` => flatten one level.
- `clone(value)`  
  Deep clone: arrays, objects and scalars are duplicated by value.
- `IsArray(value)` — returns boolean
- `IsNumeric(value)` — returns boolean (VBA-compatible `IsNumeric` semantics; `StrictIsNumeric` can be added for tighter typing).

## VBExpressions / UDF integration
- Use `@(...)` to evaluate a VBA expression during script runtime. Example:
  - Script: `return(@(ThisWBname()))`
  - Host: register UDFs in `ASF_Globals.gExprEvaluator` to expose host functions.

## VM Developer notes
- `EvalExprNode` handles evaluation of AST nodes and supports:
  - Function/closure creation: `FuncLiteral` produces closure maps containing `params`, `body`, and `env` (a ScopeStack captured by reference).
  - Call dispatch: supports builtin named calls, closures, and native VBA UDF fallbacks.
  - Member and Index evaluation: `EvalMemberNode` and `EvalIndexNode` resolve lvalues and rvalues and support assignment of members via `ResolveLValue`.

## Error handling & debugging
- The VM produces VB runtime errors for critical failures; scripts can protect runtime errors with `try { ... } catch { ... }`.
- The test harness logs `print(...)` output to `ASF_Globals.gRuntimeLog` which the `TestRunner` reads to assert behaviours.

## Examples

`flatten` usage
```js
a = [1,[2,[3]]];
print(flatten(a,1)); // -> [1,2,[3]]
print(flatten(a));   // -> [1,2,3]
```

`filter` + `reduce` chain
```js
a=[1,2,3,4,5];
return(a.filter(fun(x){ return x > 2 }).reduce(fun(acc,x){ return acc + x }, 0));
```
