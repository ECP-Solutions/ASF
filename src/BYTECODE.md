# ASF bytecode

Since version 4.0 ASF compiles every script to bytecode and runs it on a virtual machine; ASF 3
walked an AST. The classes:

| Class | Role |
|---|---|
| `ASF_Parser.cls` | The tokenizer. |
| `ASF_Compiler.cls` | Compiles a program (and the named functions it uses) from its tokens to bytecode: scope analysis, code generation, constant folding, superinstructions. |
| `ASF_VM.cls` | Runs the bytecode in a single dispatch loop, with array-based frames and an arena for closure environments. |
| `ASF_Runtime.cls` | The library of values the VM calls: methods of arrays, strings, objects, collections, Math and regex, printing, host objects and their prototypes, `@(...)`, the files of modules. |

`src/old/` holds ASF 3.2.1, the AST engine as ASF 3 users run it (ASF 3.3 and 3.4 added their
speed-ups behind conditional compilation arguments and a setting of each engine), with the suffix
`_OLD` on its classes (`ASF_OLD`, `ASF_Compiler_OLD`, `ASF_VM_OLD`...): both versions can be
imported into one project.

```vb
Dim engine As New ASF
engine.Run engine.Compile("let s = 0; for (let i = 1, i <= 100, i += 1) { s += i; }; return s;")
Debug.Print engine.OUTPUT_   ' 5050
```

## Architecture

### 0. From tokens to bytecode

`engine.Compile` hands the source to `ASF_Compiler.CompileSource`, which compiles it from its tokens
without building an AST: statements become plans of token ranges (`Array("if", condStart, condEnd,
thenStart, ...)`), and expressions are parsed and emitted in one pass. The same parser also runs
without emitting, to find the names and the function literals of a body (scope analysis), and to
find where an expression ends (the condition of `?:` compiles as jumps). When both operands of an
operator turn out to be constants, the two `PUSHK` just emitted are dropped and the folded value is
pushed; the stack depth of a function is computed at the end from the stack effect of each
instruction, so dropped instructions do not count.

Named functions keep their tokens in `gPrograms`: a later program that calls one compiles it from
there. A named function is compiled into each program that uses it.

### 1. Scope analysis at compile time

Every identifier is resolved before the program runs:

- **Globals**: the names a program assigns at its top level, `$1`, `$2`..., `Math` and injected
  variables live in a global slot array. Named functions and closures read them by slot number.
- **Locals**: parameters and the names a function assigns that no enclosing scope has get a slot in
  the function's frame.
- **Cells**: a local that a nested function literal reads or writes lives in a cell instead, so that
  the closure and the function share it. Cells are allocated in an arena owned by the VM.
- **Environment**: a function literal reads the cells it captured through its environment, a block of
  cell numbers copied when the closure is created (flat closures).

Named functions are resolved to their function number at compile time; a call to one is a direct jump.

**Scope rules.** Names are lexical: a function sees its own locals, the variables of the functions
that enclose it (for function literals) and the globals. ASF 3 looked names up at run time in the
frames of the call chain, so a named function could also read and write the locals of the function
that called it; a script that depends on that (a helper that changes a variable of its caller
without receiving it) gives a different result on ASF 4.

### 2. Bytecode in arrays

The code is three `Long` arrays, `op()`, `a()` and `b()`, plus a constant table (`Variant`). Constants
are folded at compile time, and common patterns get their own instructions
(superinstructions): `i += 1` on a local, compare and jump, `x + constant`.

### 3. A single dispatch loop

The VM dispatches with `On op GoTo` (a jump table). In LibreOffice Basic this is 3.6 times faster than
a `Select Case` over 64 cases (4.5 against 16 microseconds per instruction), and a procedure call
costs about 5 microseconds against 1.5 for the same work inline, so the frequent instructions are
written inline in the loop and only the rare ones call helpers.

Values live in a preallocated `Variant` stack. A call pushes a frame record (return address, frame
base, environment) on arrays and jumps: no VBA call, no object per call, and no VBA stack limit for
recursion.

### 4. Methods and properties without strings at run time

Method names become numbers at compile time; the VM picks the implementation by the value's
`VarType` and the method number. Property reads on objects keep a per-instruction cache of the last
position of the key and check it with one comparison before searching.

### 5. Start-up

The global scope is a slot template copied in one assignment; `Math` and the built-ins are created
once per engine.

### 6. Exceptions by table

`try/catch` regions are ranges of instructions with their handler. The dispatch loop has a single
`On Error` handler that looks up the range of the failing instruction; host calls (COM,
`CallByName`, `@(...)`) are the only other places with `On Error`.

The error rules of ASF 3 are kept: an error in a top-level statement ends the program, an error in
a function makes it return `undefined` (the caller goes on), and `throw` goes up the calls to the
nearest `catch`.

A `throw` goes up through named functions and function literals; class methods, constructors and
the callbacks of built-in methods keep it to themselves (they return `undefined`). `break` and
`continue` directly inside a `try` or a `catch`, or outside a loop, do nothing.

### Classes

Instances are objects (`type: 'Instance'`, `className`, the fields, a bound method per method), so
`typeof`, `for-in` and the properties see them as in ASF 3. Classes and methods are descriptors,
not code: a class is `{type: 'Class', name, fields, classMethods, constructor, superClass}` with the
names of its fields, a method `{type: 'Method', name, params, isStatic, fn}` and a bound method
`{type: 'BoundMethod', name, fn, instance}`, where `fn` is the number of the compiled function.
Printing an instance shows its methods that way, without their code. The constructor of each class
is compiled with the field defaults and the method bindings (the parent's fields and methods
first). `super(...)` calls the constructor of the parent of the class it is written in.

A class may extend a class of a module (imported, or exported by a module that the script imports
through another one). Each class of a module also gets two functions, kept by class number: `@init`,
which sets its fields and binds its methods on an instance (and its own parent's first), and
`@ctor`, its constructor alone. A class whose parent is not declared in its script finds them at
run time with `SUPER`, from the class value its parent's name holds: its constructor runs the
parent's `@init` before its own fields and methods, and `super(...)` the parent's `@ctor`; without
a constructor of its own, the arguments of `new` go to the parent's. Extending a value that is not
a class of a module is an error of the constructor.

### Modules and the host

- **Modules** compile to bytecode and run once per engine. Each program has its own range of
  global slots, so a module's functions read the module's globals wherever they are called from.
  The globals of modules, their classes and the cells made while loading them stay across runs,
  until `ClearModuleCache`. `export` keeps a value in the module being loaded, `import` binds it;
  `export fun` exports the function when the module ends; an exported prototype goes live on every
  import.
- **`prototype.COM...`** functions compile to bytecode; `ASF_Runtime` keeps them by host type, and a
  method call or a property read on a host object of that type runs them.
- **Methods of values** (arrays, strings, objects, collections, Math, regex, host objects) that the
  VM does not run itself go to `ASF_Runtime`, which takes the receiver and the arguments as values.
  `OverrideCollMethods` applies there. Its callbacks (`map`, `sort` with a comparator...) call back
  into the VM, whose frames go above the values in use.
- **`@(...)`** is evaluated by VBA Expressions with the variables in scope that the expression names.
- **`EnableCallTrace`**: the VM records each call with its arguments and value, and the errors
  that end a function, for `GetCallStackTrace`.

### 7. Memory

Cells and environments are numbers into arrays owned by the VM, so closures do not hold objects
that point back at them; the arrays are cleared at the end of `Run`, so the closures of a program
do not keep their scopes alive. Class instances hold a bound method per method, which points back
at its instance.

### 8. Compile errors

A source with an error raises it from `Compile`, with the message and the number (`vbObjectError +
n`) ASF 3 gave: `Expected ':' in ternary expression` (8001), `Expected ',' between object
properties, found ';'` (8004), `Rest parameter must be last parameter` (9002), unbalanced
parentheses (9024), and 9025 for the others (`Expected (condition) { ... } after if`, `Function 'f'
is already defined`, the `switch` with `case 2:`...). Every named function is parsed, also those
that no code calls. A source that does not compile leaves nothing behind: none of its functions is
registered.

## Differences from ASF 3

- Names are lexical (see *Scope rules*): a named function does not see the locals of its caller.
- `super(...)` is the parent of the class it is written in; ASF 3 took the parent of the
  instance's class, so with three levels it called the same constructor until the stack ran out.
- A class extends a class of the same script or of a module; a class declared by another script
  is not visible (each script has its own globals, as in ASF 3).
- `o.inc().inc()` runs the first `inc()` once; ASF 3 ran it twice.
- `` `${null}` `` is the text `null`, as in JavaScript; ASF 3 stopped with a type mismatch.
- `a[f()] = v` evaluates `f()` once and stores the value; ASF 3 stopped the statement.
- `(x = v);` assigns; `break` and `continue` outside a loop do nothing; any expression may be the
  step of a `for`.
- `if`, `for`, `while`, `switch`... without their parentheses or block are compile errors; ASF 3
  skipped some of them.

Kept as in ASF 3: `super.m()` gives `undefined`, and on LibreOffice `b = a; b[1] = 9` changes `a`
too, as LibreOffice Basic shares arrays between variables.

## Performance

LibreOffice 24.2, milliseconds per run of a compiled program against ASF 3.2.1 (`versionbench`,
46 workloads), and the cold time: a new engine, `Compile` and the first `Run`. Three workloads give
other results on 3.2.1, which later versions fixed: nested arrays, a `switch` with `break`, and the
value `catch` receives from `throw`.

| Workload | 3.2.1 | 4.0 | Speed-up | 3.2.1 cold | 4.0 cold |
|---|---:|---:|---:|---:|---:|
| Arithmetic, top level | 1467 | 43 | 34.5× | 1373 | 81 |
| Comparisons and branches | 1558 | 87 | 18.0× | 1884 | 82 |
| Logical operators | 2021 | 47 | 43.1× | 2147 | 85 |
| While loop | 1018 | 45 | 22.5× | 1158 | 81 |
| Function calls | 1899 | 78 | 24.5× | 2732 | 89 |
| Recursion (fib 14) | 7228 | 200 | 36.1× | 6695 | 239 |
| Deep recursion (300) | 6517 | 114 | 57.3× | 8042 | 86 |
| Closure calls | 1698 | 128 | 13.3× | 2040 | 208 |
| Array literals | 2632 | 128 | 20.5× | 3023 | 170 |
| Property reads | 2539 | 109 | 23.3× | 2364 | 151 |
| Array index write | 2359 | 501 | 4.7× | 3080 | 463 |
| String methods | 4648 | 477 | 9.7× | 5057 | 429 |
| map / filter / reduce | 1106 | 221 | 5.0× | 1051 | 291 |
| Math methods | 3272 | 165 | 19.9× | 3329 | 118 |
| Classes | 3879 | 544 | 7.1× | 4646 | 298 |
| Static methods | 3215 | 171 | 18.8× | 3130 | 224 |
| try without throw | 932 | 21 | 44.8× | 947 | 65 |
| Module function calls | 2050 | 157 | 13.1× | 1806 | 166 |
| Regex | 764 | 131 | 5.8× | 936 | 215 |
| Error in a function | 707 | 287 | 2.5× | 948 | 381 |
| Start-up (`return 1`) | 1.1 | 0.6 | 1.8× | 5.9 | 7.5 |
| Long script (300 statements) | 508 | 6 | 82.2× | 5060 | 591 |

The work `ASF_Runtime` does for the VM (strings, regex, sort with a comparator, host objects) and
errors gain least. A new engine creates a compiler and a VM besides the runtime, so `return 1`
starts slightly slower; any script with a loop or 300 statements is faster from its first run (the
300 statements, 8.6 times).

## Testing

- **Test suites**: `python3 libreoffice/tools/run_tests.py all` runs `test/TestRunner.bas` and
  `test/RegexTests.bas` on the LibreOffice edition.
- **Benchmarks**: `python3 libreoffice/tools/run_tests.py versionbench` runs every workload of
  `test/VersionBench.bas` on ASF 3.2.1 (`src/old`, loaded next to ASF 4) and ASF 4, checks the
  results and prints both times; `asfbench` runs `test/AsfBench.bas`.
- **VBA**: `python3 Tools/vba_check.py` finds the compile errors of VBA that LibreOffice lets
  through.
