
# ASF — API Reference

[![GitHub release (latest by date)](https://img.shields.io/github/v/release/ECP-Solutions/ASF?style=plastic)](https://github.com/ECP-Solutions/ASF/releases/latest) [![Tests](https://img.shields.io/badge/tests-85%2B-green.svg)](.)  

This document describes the runtime API exposed by ASF scripts and the VM builtins/methods available to scripts. It summarizes semantics, signatures, error behavior, and examples.

---

## Table of contents

- Runtime model & conventions
- Globals and program entry
- Value model
- Builtin global functions
- Array methods (exposed as properties)
- Object methods / member behavior
- [VBA Expressions](https://github.com/ECP-Solutions/VBA-Expressions) integration
- Error & truthiness rules
- Examples & usage patterns

---

## Runtime model & conventions

- **Program**: a compiled AST that can be executed by the VM. The ASF host exposes `.Compile(script)` → programIndex and `.Run(programIndex)` to run.
- **Scope / closures**: closures capture the environment by reference (shared-write semantics). That means nested functions can mutate outer-scope variables and see changes across closures.
- **Indexing base**: arrays honor `__option_base` set in runtime globals (commonly 0 or 1). All array helpers and methods handle this consistently.
- **Call signature for array callbacks**:
  - Callback receives `(value, index, array)` where `index` respects `__option_base`.
  - `this` can be supplied when calling `CallClosure` or via bound method objects (property-style access returns bound method with `baseVal` as `this`).

---

## Value model

- Core types: `Number`, `String`, `Boolean`, `Array` (Variant arrays), `Object` (`ASF_Map`), `Null` (`null` literal), `Closure` (map representing a function), `Empty`.
- Arrays are Variant arrays with explicit LBound/UBound.
- Objects are `ASF_Map` instances (Map-like behavior with `.GetValue` / `.SetValue`).

---

## Truthiness

- `IsTruthy(value)` is used internally:
  - `Falsey`: `False`, `0` (numeric zero), `""` (empty string), `Empty`, `Null`.  
  - Everything else is truthy.
  - Beware: some builtins use numeric semantics when needed.

---

## Builtin global functions

These are callable as top-level functions or via the helper bridge; some are also available as named builtins in the VM.

- `print(...)` — pretty-prints arguments and appends to runtime log (used by test suite).
- `range(limit)` or `range(start, end)` or `range(start, end, step)`  
  Returns an array of numbers.
  - Example: `range(3) // [0,1,2]`
- `flatten(array, depth?)` — returns a flattened array to `depth` (fully flattened if depth omitted).
- `clone(value)` — deep clone for arrays and objects.
- `IsArray(x)` — returns `true` if `x` is an array.
- `IsNumeric(x)` — returns `true` if `x` is numeric. 

> Many other helpers exist in the VM implementation; these are the main named builtins surfaced for scripts.

---

## Array methods (exposed as properties — e.g. `arr.map`)

All array methods are available as **properties** on arrays. Accessing `a.map` returns a bound builtin object you can call immediately: `a.map(fn)` or assign to a variable and call later `m = a.map; m(fn)`.

**Callback signature for methods that accept callbacks**: `callback(value, index, array)`  
**`this`**: For method-style calls, the base array is bound as the default `this` when the method is obtained via property access; `CallClosure` may also accept a `thisArg`.

Below is a compact table (name — signature — brief behavior — example):

- `map(callback)`  
  - Returns new array with `callback` applied to every element (does not mutate original). Supports nested arrays and closures.  
  - Example: `[1,2].map(fun(x){ return x*2 }) -> [2,4]`

- `filter(callback)`  
  - Returns new array containing elements where `callback` is truthy.  
  - Example: `[1,2,3].filter(fun(x){ return x%2==1 }) -> [1,3]`

- `reduce(callback, initial?)`  
  - Accumulates values using `callback(acc, value)`; if `initial` is omitted the first array element is used as initial (like JS).  
  - Example: `[1,2,3].reduce(fun(acc,x){ return acc+x }, 0) -> 6`

- `forEach(callback)`  
  - Calls `callback` for each element; returns `Empty` (or behavior consistent with print/log). No return collection.

- `slice(start?, end?)`  
  - Non-mutating subarray selection. Respects negative indexes and `__option_base`.

- `push(...items)`  
  - Mutates array by appending items; returns new length. Writes back to LValue container.

- `pop()`  
  - Mutates array removing last element; returns popped element. Writes back to LValue.

- `shift()`  
  - Mutates array removing first element; returns popped element. Writes back to LValue.

- `unshift(...items)`  
  - Mutates array prepending items; returns new length. Writes back to LValue.

- `concat(...itemsOrArrays)`  
  - Returns a new array concatenating the base and provided items/arrays.

- `unique()`  
  - Returns new array containing unique elements; deep-aware (structural equality for arrays/objects).

- `flatten(depth?)`  
  - Returns new flattened array up to `depth` (fully flattened if not specified).

- `clone()`  
  - Deep clone of array/object/value.

- `toString()` / `join(separator?)`  
  - Joins elements to produce a string; arrays/objects are pretty-printed for complex elements.

- `delete(index)`  
  - Remove element at user-facing index, mutate array and write back. Returns `true`/`false`.

- `splice(start, deleteCount?, ...items)`  
  - Mutating splice: removes items, inserts new ones, returns array of removed elements.

- `toSpliced(start, deleteCount?, ...items)`  
  - Non-mutating splice; returns new array with the changes applied to a copy.

- `at(index)`  
  - Returns element at index (supports negative indexing relative to end). Index respects `__option_base`.

- `copyWithin(target, start=0, end=len)`  
  - Mutates array by copying a slice to another location. Writes back to LValue.

- `entries()`  
  - Returns array of `[index, value]` pairs starting at `__option_base`.

- `every(callback)`  
  - Returns true if callback is truthy for all elements.

- `some(callback)`  
  - Returns true if callback is truthy for any element.

- `find(callback)` / `findIndex(callback)` / `findLast(callback)` / `findLastIndex(callback)`  
  - Search operations. Indices follow `__option_base`. `find` returns element or `Empty`, `findIndex` returns index or -1.

- `includes(value)` / `indexOf(value)` / `lastIndexOf(value)`  
  - Search using deep equality for complex values.

- `of(...items)`  
  - Create array from items (also exposed as global named builtin).
    
  - `from(source, mapFn?, thisArg?)` — produce an array from an array/string/single value; if `mapFn` is provided and is a closure it's applied to each element (signature `(value, index, source)`).

- `reverse()`  
  - Mutates array in place (writes back). Returns mutated array.

- `toReversed()`  
  - Non-mutating reverse; returns a new array.

- `sort(comparator?)` / `toSorted(comparator?)`  
  - `sort` mutates the array in-place and writes back. `toSorted` returns a new sorted array.
  - Comparator is a closure `fun(a,b)` returning negative/0/positive (or numeric) similar to JS.
  - Implementation uses in-place QuickSort with comparator hook.

- `with(index, value)`  
  - Returns a shallow copy with `value` set at `index` (non-mutating).

---

## Objects & members

- Objects are maps: `{ k: 1, nested: { x: 2 } }`.
- Member read: `o.x` resolves the property on the object (if base is `ASF_Map`) or if the base is an array and property matches a builtin array method name, it returns a bound builtin method (see "method/property semantics" above).
- Member write semantics use `ResolveProp` helper to ensure property writes mutate the real container (works with nested `parentObj` chains created by the compiler when producing LValue metadata).

---

## Builtin method dispatch rules

- When evaluating a `Member` node, the VM:
  1. Evaluates the `base` expression.
  2. If `base` evaluates to an `ASF_Map` (object), property lookup returns stored value (function/map/primitive).
  3. If `base` evaluates to an array and `prop` is the name of an array-method, the VM returns a `BuiltinMethod` map: `{ type: "BuiltinMethod", method: "<name>", baseVal: <array> }`.
  4. When a `Call` has a callee that is `BuiltinMethod`, the VM routes the call into the unified array-method dispatch block and executes the method with `baseValLocal` pre-bound.

This design allows `a.map(fn)` and `f = a.map; f(fn)` to behave consistently.

---

## VBA Expressions integration

- `@( ... )` syntax embeds raw VBAexpressions. The string inside is evaluated by the VBA Expressions evaluator at runtime and its result is returned to the script.
- Use-case: matrix operations, calling VBA functions/worksheet UDFs, or invoking existing code via the `gExprEvaluator` bridge.

---

## Errors & failure modes

- Many method builtins validate their arguments and return `Empty` if arguments are invalid (for example calling an array method on a non-array base will typically produce `Empty`).
- Parser/compile errors raise exceptions during `.Compile()`; runtime errors in expressions raise errors inside `.Run()` (try/catch inside ASF scripts can be used to handle runtime exceptions).
- Tests exercise many edge cases — consult `TestRunner.bas` for the canonical expected behavior.

---

## Examples

```js
// map/filter/reduce chain
a = [1,2,3,4,5];
sum = a.filter(fun(x){ return x > 2 }).reduce(fun(acc,x){ return acc + x }, 0);
return(sum); // 12

// using bound method as first-class
m = a.map;
b = m(fun(x){ return x * 10 });

// from with mapper
print(from([1,2,3], fun(x,i,arr){ return x + i })); // -> [1,3,5]

// sort with comparator
a = [3,1,2];
a.sort(fun(a,b){ return a - b });
```
