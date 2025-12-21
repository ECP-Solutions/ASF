# Changelog

All notable changes for ASF. This file combines the release notes from the project's releases.

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
