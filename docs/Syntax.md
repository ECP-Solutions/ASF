# ASF — Syntax Reference

[![GitHub release (latest by date)](https://img.shields.io/github/v/release/ECP-Solutions/ASF?style=plastic)](https://github.com/ECP-Solutions/ASF/releases/latest) [![Tests](https://img.shields.io/badge/tests-85%2B-green.svg)](.)  

This document defines the concrete syntax supported by the Advanced Scripting Framework (ASF) and contains a compact BNF grammar, operator precedence, and examples.

> Note: ASF is embedded in VBA — scripts are supplied to the ASF compiler which tokenizes and produces an AST that the VM executes. Semicolons (`;`) are used as statement separators and are required in ambiguous cases; comma (`,`) is only used as an argument or element separator.

---

## Quick summary

- Statement separator: `;`
- Argument/element separator: `,`
- Single-line comment: `/* ... */` (also supports C-style token comments in parser)
- Function literal: `fun (params...) { ... }`
- Top-level function declaration: `fun name(params...) { ... }`
- Anonymous function values are closures with **shared-write** semantics (they capture the current runtime scope by reference).
- Array literal: `[ elem1, elem2, ... ]`
- Object literal: `{ key1: value1, key2: value2 }`
- VBExpression block: `@(...)` — raw VBAexpressions block evaluated via the [VBA-Expressions](https://github.com/ECP-Solutions/VBA-Expressions) bridge.
- Null literal: `null` (literal representing absence of value)
- Boolean literals: `true`, `false`

---

## BNF grammar (compact)

This BNF uses a mixture of concrete tokens and non-terminals to show the language shape.

```
<program>        ::= <stmts>

<stmts>          ::= <stmt> ( ';' <stmt> )* [ ';' ]

<stmt>           ::= <if-stmt> | <for-stmt> | <while-stmt> | <try-stmt> | <switch-stmt> | <return-stmt> | <break-stmt> | <continue-stmt> | <expr-stmt>

<if-stmt>        ::= 'if' '(' <expr> ')' <block> ( 'elseif' '(' <expr> ')' <block> )* [ 'else' <block> ]

<for-stmt>       ::= 'for' '(' <expr> ',' <expr> ',' <expr> ')' <block>

<while-stmt>     ::= 'while' '(' <expr> ')' <block> <try-stmt>       ::= 'try' <block> 'catch' <block>

<switch-stmt>    ::= 'switch' '(' <expr> ')' '{' ( 'case' <expr> <block> )* [ 'default' <block> ] '}'

<return-stmt>    ::= 'return' [ '(' <expr> ')' | <expr> ]

<break-stmt>     ::= 'break'

<continue-stmt>  ::= 'continue'

<expr-stmt>      ::= <expr>

<block>          ::= '{' <stmts> '}' | <stmt>  -- blocks may be multiline or single statement

<expr>           ::= <ternary>

<ternary>        ::= <logical-or> [ '?' <expr> ':' <expr> ]

<logical-or>     ::= <logical-and> ( '||' <logical-and> )*

<logical-and>    ::= <bitwise-or> ( '&&' <bitwise-or> )*

<bitwise-or>     ::= <bitwise-xor> ( '|' <bitwise-xor> )*

<bitwise-xor>    ::= <bitwise-and> ( '^' <bitwise-and> )*

<bitwise-and>    ::= <equality> ( '&' <equality> )*

<equality>       ::= <relational> ( ('==' | '!=') <relational> )*

<relational>     ::= <shift> ( ('<' | '>' | '<=' | '>=') <shift> )*

<shift>          ::= <add> ( ('<<'|'>>') <add> )*

<add>            ::= <mul> ( ('+'|'-') <mul> )*

<mul>            ::= <unary> ( (''|'/'|'%') <unary> )

<unary>          ::= ('+'|'-'|'!') <unary> | <power>

<power>          ::= <postfix> ( '^' <power> )?   -- right-associative

<postfix>        ::= <primary> { <postfix-op> }*

<postfix-op>     ::= '.' IDENT                      -- member | '[' <expr> ']'                 -- index | '(' <arglist> ')'              -- call

<primary>        ::= NUMBER | STRING | 'true' | 'false' | 'null' | IDENT | '[' <elemlist> ']'             -- array literal | '{' <obj-items> '}'            -- object literal | 'fun' '(' <paramlist> ')' <block>  -- expression-level func literal | '(' <expr> ')' | '@' '(' VBA_EXPR ')'           -- VBExpr raw block

<arglist>        ::= [ <expr> ( ',' <expr> )* ]

<elemlist>       ::= [ <expr> ( ',' <expr> )* ]

<obj-items>      ::= [ (<IDENT | STRING> ':' <expr>) (',' (<IDENT|STRING> ':' <expr>))* ]

<paramlist>      ::= [ IDENT ( ',' IDENT )* ]

IDENT            ::= letter followed by letters/digits/underscore (collapsed forms allowed: e.g. "o.a[2].b" may be emitted as Ident token and expanded) NUMBER           ::= decimal or float STRING           ::= '...' or "..." VBA_EXPR         ::= any raw text until matching ')'
```

Notes:
- The parser also accepts top-level function declarations using the same `fun` syntax with a name: `fun name(params) { body }` — these are converted to program-level function definitions and stored in the global program table.
- Collapsed identifiers like `o.a[2].b` are parsed into nested AST nodes (`Variable`/`Member`/`Index`) by the compiler helper `ParseCollapsedIdentToNode`.

---

## Operator precedence & associativity

From highest precedence to lowest:

1. Parentheses `()` (grouping)
2. Postfix: calls `()`, indexes `[]`, member access `.prop` (left-to-right chaining)
3. Exponentiation `^` (right-associative)
4. Unary `+ - !`
5. Multiplicative `* / %`
6. Additive `+ -`
7. Shifts `<< >>`
8. Relational `< <= > >=`
9. Equality `== !=`
10. Bitwise/logic levels (|, ^, &, ...)
11. Logical AND `&&`
12. Logical OR `||`
13. Ternary `?:`

---

## Literals

- Number: `123`, `3.14`
- String: `'hello'` or `"hello"`
- Boolean: `true`, `false`
- Null: `null`
- Array: `[1, 2, [3], 'x']`
- Object: `{ k: 1, s: 'x' }`
- VBExpression: `@({1;0;4})` — raw block passed to VBA-expressions evaluator

---

## Functions & closures

- `fun(x,y) { return x+y }` produces a closure value.
- Top-level functions: `fun add(a,b) { return a + b }` — compiled into the global program table and callable by name.
- Closures implement **shared-write** semantics: they capture the runtime scope by reference. Mutations to outer-scope variables are visible across all closures that share that scope.

### Call semantics

- `CallClosure(closureMap, argsCollection, thisVal)`:
  - The runtime binds parameters in a new scope linked to the closure's environment.
  - Callback functions receive `(value, index, array)` when called by array-methods.
  - `this` is supported when calling closures via bound method objects or when a `thisArg` is supplied.

---

## Postfix chaining (member/call/index)

- Postfix chaining is supported for arbitrary primaries. Examples:
  - `a.b.c(d)[i].x()`
  - Special-case: `.length` on an index or array is compiled into `.__len__` builtin call at compile-time.

---

## Statement rules & semicolons

- `;` separates statements. The parser enforces semicolons more strictly to disambiguate nested constructs (recommended: terminate statements with `;` when inline or in compact code).
- Inside `{ ... }` semicolons are not required strictly before `}` but are required between adjacent statements when ambiguous.

---

## AST node types (high level)

ASF uses Map-based AST nodes internally. Common node `type` values include:

- `Literal` — { type: "Literal", value: ... }
- `Variable` — { type: "Variable", name: "x" }
- `Member` — { type: "Member", base: <node>, prop: "x" }
- `Index` — { type: "Index", base: <node>, index: <node> }
- `Call` — { type: "Call", callee: <node>, args: [<node>, ...] }
- `FuncLiteral` — { type: "FuncLiteral", params: [...], body: <stmts> }
- `Array` — { type: "Array", items: [<node>, ...] }
- `Object` — { type: "Object", items: [ (key,node), ... ] }
- `VBAexpr` — { type: "VBAexpr", expr: "..." }
- `BuiltinMethod` — runtime-built map representing a bound method (when `a.map` is evaluated and `a` is an array)

---

## Example snippets

```js
// arithmetic + precedence
return(1 + 2 * 3);

// function + closures (shared-write)
a = 1;
f = fun() { a = a + 1; return a; };
print(f());  // PRINT:2
print(a);    // PRINT:2

// arrays / map
a = [1, [2,3]];
b = a.map(fun(x) { if (IsArray(x)) { return x } else { return x * 2 } });
print(b);

// object literal & member call
o = { v: 10, incr: fun(x) { return x + 1 } };
print(o.incr(o.v));  // PRINT:11

// VBExpr embedding (evaluated by VBAexpressions)
a = @({1;0;4});
print(a);
```

Notes & hints

- null is a valid literal returned by expressions and used to represent absence of value.
- Arrays in ASF are implemented as Variant arrays and honor __option_base (runtime option that sets the base index).
- The compiler will attempt to expand collapsed identifiers like a.b[3].c into nested AST nodes so the VM can handle LValue semantics correctly.

---

References

See [TestRunner.bas](/test/TestRunner.bas) for a comprehensive test-driven specification (85+ tests that exercise syntax and runtime behavior).
