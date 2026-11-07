# Direct compiler checks (stage 3)

`ASF_Bytecode.CompileSource` (in `src/ASF bytecode/`) compiles a program from its tokens straight to
bytecode, without an AST. It must make exactly the code that the AST path makes (`ASF_Compiler`,
then `ASF_Bytecode.Prepare`); the tools of this folder check that and measure both.

## How it works

- **Plans**: a statement is a plan, an array of token ranges (`Array("if", condStart, condEnd,
  thenStart, ...)`), split as `ASF_Compiler` splits statements, quirks included. Named functions
  are registered with their tokens; classes are declarations that point at the tokens of their
  methods.
- **One parser, three modes**: emitting, scanning (the names a body uses and the function literals
  it holds, in the order the AST compiler walks them: that order numbers the closure captures) and
  probing (where an expression ends, to compile the condition of `?:` as jumps).
- **Emission**: every operand is emitted as it is parsed. When both operands of an operator are
  constants, the two `PUSHK` are dropped (with their constants, if they added them) and the folded
  value is pushed; `x + constant` becomes `ADDK` and `x = x + constant` `INCL`/`INCG`, by rewriting
  the instructions just emitted. The stack depth of a function is computed at the end from the stack
  effect of each instruction, so dropped ones do not count.
- **Token ranges** are parsed again when the order of emission needs it: the target of an
  assignment (after the value), the step of a `for` (after the body) and conditions, which compile
  as jumps from the operator at the root of the range.
- **Fallback**: what it does not take (a source with an error, `case 1:`, `export fun`) raises
  `NOT_TAKEN`, and the source goes to `ASF_Compiler`, which reports the error or builds the AST.

## Results

`python3 Tools/DirectCompiler/compare.py --sources FILE` (LibreOffice 24.2): the same disassembly on
all 571 sources (the ones the TestRunner compiles, the scripts of `test/`, 8 written by hand and 300
generated ones, half of them with functions, closures, classes, templates, `try`, `switch` and loops
over values), and on the 435 of 600 more generated programs that `ASF_Compiler` accepts. Compile
times, tokenizer included, milliseconds:

| Script | Tokenize | Tokens → AST → bytecode | Tokens → bytecode (direct) | Speed-up |
|---|---:|---:|---:|---:|
| 300 × `n = n + 1` | 169 | 1667 | 682 | 2.4× |
| 300 varied statements | 909 | 9297 | 4375 | 2.1× |

## Running it

- **LibreOffice**: `python3 Tools/DirectCompiler/compare.py` (parity, then timings); `-v` prints the
  code of any difference; `--sources FILE` adds the sources that
  `ASF_CAPTURE_SOURCES=FILE python3 libreoffice/tools/run_tests.py runner --fastvm` writes;
  `--rich N SEED` checks N generated programs only; `--export` writes `corpus.txt`, `bench1.vas`
  and `bench2.vas`.
- **Excel**: import the ASF classes (with `ASF_BYTECODE = 1`) and `DirectCompare.bas` into a workbook
  saved next to `corpus.txt`, `bench1.vas` and `bench2.vas`, and run `RunDirectComparison`; the
  results go to the Immediate window.
