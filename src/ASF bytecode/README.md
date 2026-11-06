# ASF bytecode backend

A second backend for ASF: the same language, the same front end (`ASF_Parser`, `ASF_Compiler`), and a
virtual machine that runs flat bytecode instead of walking the AST. The classes of `src/` stay as they
are; this folder adds two:

| Class | Role |
|---|---|
| `ASF_Bytecode.cls` | Compiles the AST of a program (and the named functions it uses) to bytecode: scope analysis, code generation, constant folding, superinstructions. |
| `ASF_FastVM.cls` | Runs the bytecode in a single dispatch loop, with array-based frames and an arena for closure environments. |

## Using it

1. Import the classes of `src/` and the two classes of this folder.
2. Set the conditional compilation argument `ASF_BYTECODE = 1` (VBA editor: Tools > VBAProject
   Properties > Conditional Compilation Arguments). Without it, `ASF.cls` compiles without these classes.
3. Turn the backend on per engine:

```vb
Dim engine As New ASF
engine.FastVM = True
engine.Run engine.Compile("let s = 0; for (let i = 1, i <= 100, i += 1) { s += i; }; return s;")
Debug.Print engine.OUTPUT_, engine.Backend   ' 5050  bytecode
```

A program that uses something the bytecode compiler does not take yet runs whole on the tree VM;
`engine.Backend` says which one ran it and `engine.FallbackReason` why. With the argument
`ASF_FASTVM_DEFAULT = 1` as well (`ASF_BYTECODE = 1 : ASF_FASTVM_DEFAULT = 1`), every engine starts
with `FastVM` on, which runs a whole test suite on the bytecode backend. The LibreOffice extension
(`ASF.oxt`) includes both backends.

## Architecture

### 1. Scope analysis at compile time

Every identifier is resolved before the program runs:

- **Globals**: the names a program assigns at its top level, `$1`, `$2`..., `Math` and injected
  variables live in a global slot array. Named functions and closures read them by slot number.
- **Locals**: parameters and the names a function assigns that no enclosing scope has get a slot in
  the function's frame.
- **Cells**: a local that a nested function literal reads or writes lives in a cell instead, so that
  the closure and the function share it, as the tree VM's shared frames do. Cells are allocated in an
  arena owned by the VM.
- **Environment**: a function literal reads the cells it captured through its environment, a block of
  cell numbers copied when the closure is created (flat closures).

Named functions are resolved to their function number at compile time; a call to one is a direct jump.

**Scope rules.** The bytecode VM resolves names lexically: a function sees its own locals, the
variables of the functions that enclose it (for function literals) and the globals. The tree VM looks
names up at run time in the frames of the call chain, so a named function there can also read and
write the locals of the function that called it. Programs that depend on that (a helper that changes
a variable of its caller without receiving it) give different results; the differential tests list
the ones found.

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

The tree VM's error rules are kept: an error in a top-level statement ends the program, an error in
a function makes it return `undefined` (the caller goes on), and `throw` goes up the calls to the
nearest `catch`.

A `throw` goes up through named functions and function literals, as in the tree VM; class methods,
constructors and the callbacks of built-in methods keep it to themselves (they return `undefined`).
`break` and `continue` directly inside a `try` or a `catch` do nothing, as there.

### Classes

Instances keep the tree VM's layout (`type: 'Instance'`, `className`, the fields, a bound method per
method), so `typeof`, `for-in` and the properties give the same results. Classes and methods are
descriptors, not code: a class is `{type: 'Class', name, fields, classMethods, constructor,
superClass}` with the names of its fields, a method `{type: 'Method', name, params, isStatic, fn}`
and a bound method `{type: 'BoundMethod', name, fn, instance}`, where `fn` is the number of the
compiled function. Printing an instance shows its methods that way, without their code. The
constructor of each class is compiled with the field defaults and the method bindings of
`InstantiateClass` (the parent's fields and methods first). `super(...)` calls the constructor of the parent of the class it is written in; the
tree VM looks the parent up from the class of the instance, which recurses with three levels.

### Modules and the host

- **Modules** compile to bytecode and run once per engine, as in the tree VM. Each program has its own
  range of global slots, so a module's functions read the module's globals wherever they are called
  from. The globals of modules, their classes and the cells made while loading them stay across
  runs, until `ClearModuleCache`. `export` keeps a value in the module being loaded, `import` binds
  it; an exported prototype goes live on every import.
- **`prototype.COM...`** functions compile to bytecode; the tree VM's registry keeps the function,
  and a method call or a property read on a host object of that type runs it.
- **Methods of values** (arrays, strings, objects, collections, Math, regex, host objects) that the
  bytecode VM does not run itself go to the method library of `ASF_VM`, which takes the receiver and
  the arguments as values: no AST is built for them. `OverrideCollMethods` applies there.
- **`@(...)`** is evaluated by VBA Expressions with the variables in scope that the expression names.

No AST is used while a program runs: the constant table holds values and descriptors only.

With `EnableCallTrace` a program runs on the tree VM, so that the trace shows its calls.

### 7. Memory

Cells and environments are numbers into arrays owned by the VM, so closures do not hold objects that
point back at them; the arrays are cleared at the end of `Run`, so the closures of a program no
longer keep their scopes alive as the tree VM's do. Class instances keep the tree VM's layout, in
which each bound method points back at its instance.

## Phases

| Phase | Content | Status |
|---|---|---|
| 0 | Engine option and fallback, LibreOffice build, differential test harness, benchmarks | done |
| 1 | Expressions, variables, `if`/`while`/`for`, named functions | done |
| 2 | Closures with cells, arrays, objects with caching | done |
| 3 | Built-in methods, templates, `switch`, `for-of`/`for-in`, spread/rest, destructuring | done |
| 4 | Classes, `this`/`super`, `try/catch` | done |
| 5 | Modules, `prototype` on COM classes, `OverrideCollMethods`, `@(...)` | done |

## Differences from the tree VM

Found by the differential tests; the bytecode VM gives the first result:

- Names are lexical (see *Scope rules*): a named function does not see the locals of its caller.
- `super(...)` is the parent of the class it is written in; the tree VM takes the parent of the
  instance's class, so with three levels it calls the same constructor until the stack runs out.
- `o.inc().inc()` runs the first `inc()` once; the tree VM runs it twice.
- `` `${null}` `` is the text `null`, as in JavaScript; the tree VM stops with a type mismatch.

Kept as in the tree VM: `super.m()` gives `undefined`, and on LibreOffice `b = a; b[1] = 9`
changes `a` too, as LibreOffice Basic shares arrays between variables.

## Performance

LibreOffice 24.2, milliseconds per run of a compiled program (`fastbench`, 46 workloads; all of them
run on the bytecode VM and give the tree VM's results):

| Workload | Tree | Bytecode | Speed-up |
|---|---:|---:|---:|
| Arithmetic, top level | 652 | 41 | 15.8× |
| While loop | 649 | 31 | 21.1× |
| Function calls | 1090 | 72 | 15.2× |
| Recursion (fib 14) | 3240 | 248 | 13.1× |
| Deep recursion (300) | 2484 | 60 | 41.7× |
| Closure calls | 1032 | 125 | 8.3× |
| Property reads | 1141 | 113 | 10.1× |
| Array index write | 1355 | 348 | 3.9× |
| Nested arrays | 1977 | 807 | 2.4× |
| String methods | 1633 | 348 | 4.7× |
| map / filter / reduce | 540 | 127 | 4.2× |
| Math methods | 1187 | 100 | 11.9× |
| Classes | 1503 | 174 | 8.6× |
| Static methods | 1367 | 133 | 10.3× |
| try without throw | 574 | 21 | 27.1× |
| throw and catch | 250 | 19 | 12.9× |
| Module function calls | 1071 | 516 | 2.1× |
| Regex | 425 | 225 | 1.9× |

Work done by the tree VM for the bytecode VM (regex, sort with a comparator, host objects, module
functions) gains least. The bytecode compiler runs once per program: on a 300-statement script the
first run takes about 20% longer than on the tree VM, and every later run is 50 times faster.
`asfbench` (the workloads of `test/AsfBench.bas`) gives 2.7× (nested arrays) to 15.6× (recursion).

## Testing

- **Differential tests**: `python3 libreoffice/tools/run_tests.py runner --fastvm` runs the whole
  TestRunner suite with every engine on the bytecode backend, and reports which programs fell back
  to the tree VM and why.
- **Benchmarks**: `python3 libreoffice/tools/run_tests.py fastbench` runs every workload of
  `test/FastVMBench.bas` on both backends, checks that the results agree and prints both times;
  `asfbench --fastvm` runs `test/AsfBench.bas` on the bytecode VM.
