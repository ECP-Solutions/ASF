---
layout: default
title: Syntax
has_children: false
nav_order: 3
description: "ASF Public API."
---

# ASF — API Reference

[![GitHub release (latest by date)](https://img.shields.io/github/v/release/ECP-Solutions/ASF?style=plastic)](https://github.com/ECP-Solutions/ASF/releases/latest) [![Tests](https://img.shields.io/badge/tests-290%2B-green.svg)](.)  

This document describes the runtime API exposed by ASF scripts and the VM builtins/methods available to scripts. It summarizes semantics, signatures, error behavior, and examples.

---

## Table of contents

- Runtime model & conventions
- Globals and program entry
- Value model
- Builtin global functions
- Array and String methods (exposed as properties)
- Regex Object 
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
- `forEach(arr, fn)` — executes the provided function once for each array element. This method returns the original array, doesn't behave like the javascript alternative.
- `Regex(pattern?, ignoreCase?, multiline?, dotAll?)` — regex Object constructor. If all arguments are omitted, the regex Object will not initialized. If one argument is given, this is considered the `pattern`; if two arguments are given, `pattern` and `ignoreCase` flag will be set; if three arguments are given, `pattern`, `ignoreCase` and `multiline` flags will be set; when all arguments are givem, also the `dotAll` is set. 

> Those are the current named builtins surfaced for scripts.

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

## string methods (exposed as properties — e.g. `str.toLowercase`)

All string indexing in ASF is **zero-based**. String methods that accept negative indexes treat them as offset from the end (like JavaScript `.slice()` / `.at()`).

> Templates and regex patterns must be enclosed by the backtick character "\`".

### Instance methods (call on a string value, e.g. 'hello'.method(...))
- `length` — `str.length` — Number of UTF-16 code units in the string (like JS `length`). — `'abc'.length -> 3`

- `at(index)` — `str.at(index)` — Return character at `index`. Supports negative indexes (e.g. `-1` is last char). Returns `''` for out-of-range. — `'abc'.at(1) -> 'b'`, `'abc'.at(-1) -> 'c'`

- `charAt(index)` — `str.charAt(index)` — Same as `at(index)` (returns single-character string or `''`). — `'abc'.charAt(2) -> 'c'`

- `charCodeAt(index)` — `str.charCodeAt(index)` — Returns numeric UTF-16 code of character at `index`. If out-of-range returns `Empty` (or implementation-defined numeric error). — `'A'.charCodeAt(0) -> 65`

- `concat(...items)` — `str.concat(item1, item2, ...)` — Return new string by concatenating `str` with provided values (non-strings are coerced). Does **not** mutate. — `'a'.concat('b', 3) -> 'ab3'`

- `endsWith(search, pos?)` — `str.endsWith(search[, length])` — Tests whether string ends with `search` (optionally considering only up to `length` characters). Returns boolean. — `'abc'.endsWith('c') -> True`

- `includes(sub)` — `str.includes(sub)` — `True` if `sub` appears anywhere in `str`. — `'abc'.includes('b') -> True`

- `indexOf(search, pos?)` — `str.indexOf(search[, fromIndex])` — Index of first occurrence; returns `-1` when not found. Negative `fromIndex` treated as `0`. — `'banana'.indexOf('na') -> 2`

- `lastIndexOf(search, pos?)` — `str.lastIndexOf(search[, fromIndex])` — Index of last occurrence <= `fromIndex`. Returns `-1` if none. — `'banana'.lastIndexOf('na') -> 4`

- `localeCompare(other)` — `str.localeCompare(other)` — simple VBA based implementation, does not behave like javascript: returns negative, 0, or positive number like JS. — `'a'.localeCompare('b') -> -1`

- `match(patternOrSlashRegex)` — `str.match(pattern)` — Works two ways:
  - If `pattern` is a slash-regex string `'/pat/flags'` or a regex object, behaves like JS `match`.
    - If `g` flag **present** (global), returns an array of **strings** (just the matches; groups ignored).
    - If `g` flag **absent**, returns an array-like collection: `[fullMatch, group1, group2, ...]` (or `Empty` if no match).
  - If `pattern` is a plain string, the first occurrence is returned in the array as full match (no groups).
  - Example: 
			```js
			'a1b2'.match(`/\\d/`) /*-> ['1', ...groups? none]*/; 'a1b2'.match(`/\\d/g`) //-> ['1','2']
			```

- `matchAll(patternOrSlashRegex)` — `str.matchAll(pattern)` — JS-like behavior:
  - **Requires** the `g` (global) flag on regex (throw/raise error at runtime if omitted).
  - Returns an array of matches (strings) for global matches (no groups). If you need groups, use `ExecAll` or `regex` object methods.
  - Example: 
			```js
			'a1b2'.matchAll(`/\\d/g`) /*-> ['1','2']*/; 'a1b2'.matchAll(`/\\d/`) // -> **error**
			```

- `padEnd(targetLength, padString?)` — `str.padEnd(len[, pad])` — Pads on the right to `len` using `pad` (defaults to space). Returns new string. — `'a'.padEnd(3,'-') -> 'a--'`

- `padStart(targetLength, padString?)` — `str.padStart(len[, pad])` — Pads on the left. — `'a'.padStart(3,'0') -> '00a'`

- `repeat(count)` — `str.repeat(count)` — Returns repeated string `count` times. `count` must be non-negative integer. — `'ab'.repeat(2) -> 'abab'`

- `replace(searchValue, replaceValue)` — `str.replace(search, repl)` — Replaces **first** match when `search` is a string, or behavior depends on regex flags:
  - `search` may be a plain string (literal) or a slash-regex string `'/pattern/flags'` or a regex object.
  - `replaceValue` may be a string or a function/closure. If string, supports `$`-expansions:
    - `$&` — full match
    - `$n` — numbered capture `$1..$99`
    - ``$` `` — left context (text before match)
    - `$'` — right context (text after match)
    - `$$` — literal `$`
  - If `replaceValue` is a closure/map, it is invoked as `callback(match, p1, p2, ..., offset, originalString)` and its return coerced to string (complex values pretty-printed).
  - `offset` is **zero-based** (consistent with JavaScript).
  - Example: 
			```js
			'a1b'.replace(`/(\\d)/`, fun(m,p,off,s){ return '[' & p & ']'}) // -> 'a[1]b'
			```

- `replaceAll(searchValue, replaceValue)` — `str.replaceAll(search, repl)` — Same as `replace` but replaces **all** matches. Accepts slash-regex `g` or will do global if called as `replaceAll` (or if regex flag `g` present). Function replacement, `$` expansions supported. Safeguards for zero-length matches (engine advances one char to avoid infinite loops).

- `slice(start?, end?)` — `str.slice(start?, end?)` — Extracts substring from `start` up to (but not including) `end`. Supports negative indices counted from end. Non-mutating. — `'abcdef'.slice(1,4) -> 'bcd'`

- `split(separator?, limit?)` — `str.split(sep, limit?)` — If `sep` is a slash-regex string or regex object, uses the regex engine (honors `g`/flags). If `sep` is `''`, split into characters. Returns array. `limit` optional. — `'a,b;c'.split('/[;,]/') -> ['a','b','c']`

- `startsWith(search, pos?)` — `str.startsWith(search[, pos])` — Tests whether `str` starts with `search` at optional `pos`. Returns boolean. — `'abc'.startsWith('b',1) -> True`

- `substring(start?, end?)` — `str.substring(start?, end?)` — Returns substring between `start` and `end`. Negative values are treated as `0`. If `start> end` they are swapped (JS behavior). — `'abcd'.substring(2,1) -> 'bc'`

- `toLowercase()` — `str.toLowercase()` — Return lowercased string. (Name: `toLowercase` per ASF naming.) — `'ABC'.toLowercase() -> 'abc'`

- `toUppercase()` — `str.toUppercase()` — Return uppercased string. — `'abc'.toUppercase() -> 'ABC'`

- `trim()` — `str.trim()` — Removes whitespace from both ends. — `'  hi  '.trim() -> 'hi'`

- `trimStart()` — `str.trimStart()` — Removes leading whitespace. — `'  hi'.trimStart() -> 'hi'`

- `trimEnd()` — `str.trimEnd()` — Removes trailing whitespace. — `'hi  '.trimEnd() -> 'hi'`

- `toString()` — `str.toString()` — Identity/primitive string representation. — `('x').toString() -> 'x'`

- `valueOf()` — `str.valueOf()` — Returns primitive string value (alias of `toString`). — `'x'.valueOf() -> 'x'`

---

## Regex Object API
ASF provides a **native regex engine** exposed through the `regex()` constructor and slash-regex syntax.
  
Regex objects are **stateful**, configurable, and reusable, and integrate seamlessly with string methods like `match`, `replace`, `split`, etc.

### Execution & Matching
- `exec(subject)`
	- `regex.exec(subject)`
	- Executes the regex on `subject` and returns **full match + capture groups** as an array.
	- Returns `Empty` if no match.
	- Equivalent to JS `RegExp.prototype.exec` (non-global).
	- Example:
		```js
		regex(`(a)(b)`).exec('xab')
		// -> ['ab', 'a', 'b']
		```
- `execAt(subject, position)`
	- `regex.execAt(subject, pos)`
	- Attempts a match starting **exactly at** `pos` (1-based index).
	- Returns full match + groups or `Empty`.
	- Useful for scanners and parsers.
	- Example:
		```js
		r = regex(`\\d+`)
		r.execAt('a123', 2)
		// -> ['123']
		```
- `execAll(subject)`
	- `regex.execAll(subject)`
	- Returns **all matches** as an array.
	- Each element is the **full match string only** (groups ignored).
	- Equivalent to JS global matching.
	- Example:
		```js
		regex(`\\d+`).execAll('a1b22c333')
		// -> ['1', '22', '333']
		```
- `test(subject)`
	- `regex.test(subject)`
	- Returns `True` if the regex matches anywhere in `subject`, `False` otherwise.
	- Example:
		```js
		regex(`\d`).test('abc1')
		// -> True
		```
### Replacement & Splitting
- `replace(subject, replacement)`
	- `regex.replace(subject, replacement)`
	- Replaces **all matches** in `subject`.
	- `replacement` is a string.
	- Supports `$` expansions:
		- `$0` full match
		- `$1..$99` capture groups
	- Example:
		```js
		regex(`(\d+)`).replace('a12b', '[$1]')
		// -> 'a[12]b'
		```
- `split(subject)`
	- `regex.split(subject)`
	- Splits the string using the regex as delimiter.
	- Returns an array of substrings.
	- Safely handles zero-length matches.
	- Example:
		```js
		regex(`[,;\s]+`).split('apple,orange;banana grape')
		// -> ['apple','orange','banana','grape']
		```
### Utility
- `escape(text)`
	- `regex.escape(text)`
	- Escapes **all regex meta-characters** in `text`.
	- Safe for building literal patterns dynamically.
	- Example:
		```js
		regex().escape('a+b*c?')
		// -> '\a\+b\*c\?'
		```
### Configuration Methods (Getters / Setters)
Regex objects are **mutable** and configurable at runtime.

**Pattern**
- `getPattern()`
	- Returns the current regex pattern string.
	- Example: `r.getPattern()`
- `setPattern(pattern)`
	- Sets a new pattern (does not reset flags).
	- Example: 
		```js
		r.setPattern(`\d+`)
		```
**Flags**
- `getIgnoreCase()` / `setIgnoreCase(value)`
	- Controls case-insensitive matching (`i` flag).
	- Example:
		```js
		r.setIgnoreCase(true)
		```
- `getMultiline()` / `setMultiline(value)`
	- Enables multiline mode (`m` flag).
	- `^` and `$` match line boundaries.
	- Example:
		```js
		r.setMultiline(true)
		```
- `getDotAll()` / `setDotAll(value)`
	- Enables dotAll mode (`s` flag).
	- `.` matches newline characters.
	- Example:
		```js
		r.setDotAll(true)
		```
**Execution Control**
- `getMaxMatchSteps()` / `setMaxMatchSteps(value)`
	- Limits internal backtracking steps.
	- Prevents catastrophic regex behavior.
	- Example:
		```js
		r.setMaxMatchSteps(50000)
		```
**Initialization Helper**
- `init(pattern, ignoreCase?, multiline?, dotAll?)`
	- Reinitializes the regex object safely.
	- Returns `True` on success, `False` on error.
	- Example:
		```js
		r.init(`\w+`, true, true, false)
		```
---

## String Templates

- Templates are an special type of strings holding expressions and literals: `<literal>${<expression>}<literal>?`.
- Each `${...}` block is captured as an expression exactly as typed respecting escaping rules inside.
- The tokenizer allows escape this chars for string Templates: "\`", "/", "\\", "$", "{", "}"

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

// Advanced replacements
fun replacer(match, p1, p2, p3, offset, string)
	{return [p1, p2, p3].join(' - ');};
newString = 'abc12345#$*%'.replace(`/(\D*)(\d*)(\W*)/`, replacer);
print(newString);
```
