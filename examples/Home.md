# Advanced Scripting Framework (ASF) Language Reference And Examples
[![Tests (Rubberduck)](https://img.shields.io/badge/tests-Rubberduck-brightgreen)](https://rubberduckvba.com/)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
[![GitHub release (latest by date)](https://img.shields.io/github/v/release/ECP-Solutions/ASF?style=plastic)](https://github.com/ECP-Solutions/ASF/releases/latest)

Below is a single-file reference describing the **ASF** scripting language features, usage examples, and the exact syntax rules the compiler/VM expect.

> **Important**: This document follows this ASF compiler/runtime:
> - `;` is the **statement separator** (mandatory to end statements).
> - `,` is **only** for separating function arguments, array items, and object members.
> - The engine supports: literals, variables, one top named `fun name(...) { ... }`, expression-level function literals (`fun(...) { ... }`), closures (shared-write semantics), `if`, `for` / `while`, `try/catch`, `return`, `break` / `continue`, arrays `[...]`, objects `{...}`, templates `` `...${...}...` ``, slash-regex strings `"/pattern/flags"` and a `regex()` constructor.

---

## Table of contents

1. Quick syntax rules  
2. Values and literals  
3. Expressions and operators  
4. Statements and required statement terminator (`;`)  
5. Functions & closures (top-level limitation)  
6. Control flow: `if`, `switch`, `for`, `while`, `break`/`continue`  
7. Try / Catch  
8. Arrays and array methods — `[]` and `.method` on `[]` literal forms  
9. Objects: `{ key: value, ... }` and `.method` on `{}` literal forms  
10. Strings, template literals, and string methods (JS-like)  
11. Regular expressions — two flavors and `regex()` constructor  
12. Replacement string `$` expansion in `replace` / `replaceAll`  
13. Method chaining and temporary (literal) receivers  
14. Escaping rules summary (templates & regex)  
15. Error / unsupported features & common pitfalls  
16. Appendix — examples for common patterns

---

## 1. Quick syntax rules

- **Statements must end with a semicolon `;`**. Omitting `;` will make the compiler treat tokens as part of the following statement.
- **Commas `,` are exclusively for**:
  - function call arguments: `f(a, b, c);`
  - array items: `[1, 2, 3]`
  - object member separators: `{ k1: v1, k2: v2 }`
- Arrays and objects are created with `[...]` and `{...}` respectively.

---

## 2. Values and literals

- **Numeric literal**: `123`, `3.14`  
- **String literal**: `'hello'`  
- **Template literal**: backtick form with embedded expressions: `` `hello ${name}` ``  
- **Array literal**: `[1, 'a', 3]`  
- **Object literal**: `{ key: 'value', foo: 1 }`  
- **Function literal** (expression-level): `fun(x, y) { return x + y; }`  
- **Top-level named function**: `fun main(a,b) { ... }` — only one top-level named function is accepted; other functions should be function literals assigned to variables.

---

## 3. Expressions and operators

- Arithmetic: `+`, `-`, `*`, `/`, `%`, `^` (exponent right-associative)  
- Comparison: `==`, `!=`, `<`, `<=`, `>`, `>=`  
- Logical: `&&`, `||`, `!`  
- Bitwise (if implemented): `&`, `|`, `^`, `<<`, `>>`  
- String concatenation: `+` (JS-like)  
- **Ternary**: `cond ? exprTrue : exprFalse` (right-associative)

---

## 4. Statements

- `if`, `else`:
  ```js
  if (cond) { ... }
  else { ... }
  ```
- `return` to return from a named or anonymous function.
- **Statement terminator**: **every** statement must end with `;`.

---

## 5. Functions and closures

- **Named top-level function**:
  ```js
  fun main() {
    print("top-level main");
  }
  ```
  Only **one** named top-level `fun` is supported by the program loader.

- **Function literal**:
  ```js
  f = fun(x, y) { return x + y; };
  ```

- **Closures** capture the runtime scope by reference (shared-write semantics). Mutations inside closures reflect in the outer scope.

- **Replace callback signature**:
  `function replacer(match, p1, p2, ..., offset, originalString)`

---

## 6. Control flow: for / while / break / continue

- C-like `for`:
  ```js
  for (i = 0; i < 10; i = i + 1) {
      ...
  }
  ```
- `while`:
  ```js
  while (cond) {
      ...
  }
  ```
- `break` and `continue` supported.

---

## 7. try / catch

- Basic usage:
  ```js
  try {
    risky();
  } catch (e) {
    print("err: " + e);
  }
  ```

---

## 8. Arrays and methods

- Literal: `[1, 2, 3]`  
- Access: `arr[1]` (ASF actually use 1-based arrays)
- Allowed temporary literal method calls: `[].method(...)`

---

## 9. Objects and methods

- Literal: `{ a: 1, b: 'x' }`  
- Access: `obj.a` or `obj["a"]`  
- Temporary literal method calls: `({a:1}).f();`

---

## 10. Strings, template literals, and string methods

### Template literal rules
- Syntax: backtick-delimited with `${...}` expressions:
  ```js
  `hello ${ name } world ${ 1 + 2 }`;
  ```
- **Tokenizer rules** (engine design):
  - Outside `${...}`: a backslash `\` escapes only: ``\` ``, `\/`, `\\`.
  - Inside `${...}`: **also** — escapes `$`, `{`, `}`.
  - Nested `${...}` is supported. The tokenizer emits a sequence of `LITERAL` and `EXPR` parts; nested placeholders appear as additional `EXPR` parts (split behavior).
- Compile-time: each `EXPR` part is parsed into an AST; each `LITERAL` is stored as a string node.
- VM-time: parts are evaluated left-to-right and concatenated.

### String methods
- `.replace(searchValue, replaceValue)` — replaces the **first** occurrence.
- `.replaceAll(searchValue, replaceValue)` — replaces **all** occurrences.
  - `searchValue` can be a plain string or a slash-regex `"/pattern/flags"`.
  - `replaceValue` can be a string with `$` expansions, or a function (closure).
- Replacement function will be called as: `(match, p1, p2, ..., offset, originalString)`.

---

## 11. Regular expressions

Two ways:

1. **Slash-regex string**: "\`/pattern/flags\`" used inline when calling string methods. Flags: `g`, `i`, `m`, `s`.  
2. **regex() constructor**: r = regex(\`a+b\`) then configure properties or call methods.
	- r = regex(\`a+b\`, True) // set the ignoreCase flag
	- r = regex(\`a+b\`, True, True) // set the ignoreCase and multiline flags
	- r = regex(\`a+b\`, True, True, True) // set the ignoreCase, multiline and dotAll flags
Supported features:
- Classes, quantifiers (greedy/lazy/possessive), atomic groups `(?>...)`, lookahead, **fixed-width** lookbehind, alternation, anchors.
Unsupported or restricted:
- Variable-length lookbehind -> compile-time error.
- Some advanced features like backreferences or `\p{...}`.

---

## 12. Replacement `$` expansions

String replacement supports:
- `$0` — full match
- `$1..$n` — capture groups
- `$$` — literal `$` (if implemented)
If replacement is a function, its return value is stringified.

---

## 13. Method chaining and temporary receivers

You may call `.method()` directly on literals:
```js
' a '.trim().split(' ');
[1,2,3].map(fun(x){return x+1;}).join(',');
({x:1}).f();
```

---

## 14. Escaping summary

### Template literals
- Outside `${}`: `\` escapes only `` ` ``, `/`, `\`.
- Inside `${}`:  **also** — escapes `$`, `{`, `}`.

### Slash-regex strings
- Use \`/pattern/flags\` and escape regex metacharacters inside the `pattern` with `\`.

---

## 15. Errors & common pitfalls

- Missing `;` leads to parse issues.
- Only one top-level `fun name(...) {}`.
- Variable-length lookbehind causes compile error.
- Inline flags without `:` like `(?i)` may be unsupported — use `(?i:...)` or regex object setup.

---

## 16. Appendix — examples

### Basic
```js
a = 10;
b = a * 2;
print(b);
```

### Function & closure
```js
fun main() {
  makeCounter = fun() {
    n = 0;
    return fun() { n = n + 1; return n; };
  };
  c = makeCounter();
  print(c()); // 1
  print(c()); // 2
}
```

### Template
```js
name = 'Bob';
`Hello ${ name.toUpperCase() }, you have ${ 3 + 2 } messages.`;
escaped = `literal with backtick \` and slash \/ and backslash \\`;
nested = `outer ${ inner ${ a } } tail`;
```

### Replace with function
```js
fun replacer(match, p1, p2, p3, offset, s) {
  return [p1, p2, p3].join(' - ');
}
newString = 'abc12345#$*%'.replace(`/(\\D*)(\\d*)(\\W*)/g`, replacer);
print(newString);
```

### Regex object
```js
r = regex(`a+b`);
r.setignorecase(true);
caps = r.exec('aaab');
if (caps) {
  print(caps.item(1));
}
```

---