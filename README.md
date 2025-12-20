# Advanced Scripting Framework (ASF)

... (keep existing README top-level content if present) ...

## Language additions

### typeof operator

The `typeof` operator returns a string describing the runtime type of a value. Example:

```js
f = fun(a) { a = a + 1; return a };
print(typeof f);        // -> 'function'
print(typeof f(1));     // -> 'number'
```

### for-in and for-of loops

ASF supports both `for-in` and `for-of` loops.

- `for (i in a) { ... }` — iterates over keys/indices of the iterable `a`.
  - For arrays, indices reflect the current `Option Base` (1 or 0).
  - For maps/objects, keys are iterated in insertion order when the map preserves insertion order; otherwise iteration may be in sorted or implementation-defined order.

- `for (v of a) { ... }` — iterates over values of the iterable `a`.
  - For arrays, yields each element value in index order.
  - For maps/objects, yields values in the same ordering as `for-in` over keys.

Examples and expected behavior:

- for-in array (indices reflect option base)

```js
// if Option Base = 1
GetResult "a=[10,20]; out=[]; for (i in a) { out.push(i) }; print(out);"
// expect: [1,2]
```

- for-of array (values)

```js
GetResult "a=[1,2,3]; sum=0; for (v of a) { sum=sum+v }; return(sum);"
// expect: 6
```

- for-in map (insertion order or sorted if insertion order not exposed)

```js
GetResult "m = {a:1, b:2, c:3}; out=[]; for (k in m) { out.push(k) }; print(out);"
// expect: ['a','b','c'] (in insertion order or sorted if map doesn't expose insertion order)
```

- for-of map (values)

```js
GetResult "m = {a:1, b:2}; vals=[]; for (v of m) { vals.push(v) }; print(vals);"
// expect: [1,2]
```

---
