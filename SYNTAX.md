# ASF Syntax Reference

... (keep existing syntax sections if present) ...

## typeof

Syntax:

```
typeof Expression
```

Semantics:
- Evaluates Expression and returns a string representing its runtime type. Typical return values include: `number`, `string`, `boolean`, `function`, `object`, `null`, `undefined` (if applicable), etc.

Examples:

```js
typeof 1            // 'number'
typeof 'x'          // 'string'
typeof fun() { }    // 'function'
```

## for-in

Syntax:

```
for (Identifier in Expression) { StatementList }
```

Semantics:
- If Expression evaluates to an array, iteration yields indices. The numeric indices produced follow the environment's `Option Base` setting (0 or 1). Example: with `Option Base = 1`, iterating over `[10,20]` with `for (i in a)` yields `i` values `1` then `2`.
- If Expression evaluates to a map/object, iteration yields keys. The ordering is the map's iteration ordering: insertion order when the map preserves it; otherwise an implementation-defined order (commonly sorted or insertion-approximate order).

Examples:

```js
// Arrays: indices reflect Option Base
// Option Base = 1
a = [10,20]
out = []
for (i in a) { out.push(i) }
// out -> [1,2]
```

```js
// Maps/objects: keys in insertion order (if available)
m = {a:1, b:2, c:3}
out = []
for (k in m) { out.push(k) }
// out -> ['a','b','c']
```

## for-of

Syntax:

```
for (Identifier of Expression) { StatementList }
```

Semantics:
- If Expression evaluates to an array, iteration yields element values in index order.
- If Expression evaluates to a map/object, iteration yields values in the same sequence corresponding to the map's key iteration order.

Examples:

```js
// Arrays: values
a = [1,2,3]
sum = 0
for (v of a) { sum = sum + v }
// sum -> 6
```

```js
// Maps: values
m = {a:1, b:2}
vals = []
for (v of m) { vals.push(v) }
// vals -> [1,2]
```

---
