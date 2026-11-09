# Direct compiler prototype (stage 3)

`ASF_Direct.cls` compiles a program from its tokens straight to the bytecode of
`src/ASF bytecode/ASF_Bytecode.cls`, in one pass and without an AST. It is the prototype of stage 3
of the plan to retire the AST: it takes a subset of the language, and it exists to measure the gain
and to prove that a direct compiler can make exactly the code the AST compiler makes.

## What it takes

Top-level code: assignments (`=`, `+=`, `-=`, `*=`, `/=`, to a variable, `x.p`, `x[i]` and chains
of them), expression statements, `if`/`elseif`/`else`, `while`, `for (init, cond, step)`, `return`,
`print`; numbers, strings, `true`, `false`, `null`, `undefined`, variables, array and object
literals, `.prop`, `[index]` (also the collapsed `a[1]` tokens of the tokenizer), unary `-`, `!`,
`typeof`, every binary operator, `&&`, `||`, `??` and `?:`. Anything else (functions, calls,
classes, templates...) makes `Compile` return `False` and `Reason` say why.

## How it works

- **Globals**: one scan of the tokens finds the names assigned at the start of a statement, in
  order: they are the first global slots, as `ASF_Bytecode.CollectAssigned` numbers them.
- **Emission**: every operand is emitted when it is parsed. When both operands of an operator turn
  out to be constants, the two `PUSHK` are dropped (with their constants, if they were the last
  ones added) and the folded value is pushed; `x + constant` becomes `ADDK`, and `x = x + constant`
  becomes `INCG`, by rewriting the instructions just emitted. The stack depth is computed at the end
  from each instruction's stack effect, so dropped instructions do not count.
- **Token ranges**: what the order of emission needs later is a token range parsed again: the target
  of an assignment (emitted after the value), the step of a `for` (after the body), and conditions,
  which compile as jumps (`JNLT`..., `&&`/`||` as chains of jumps) from the range: the root operator
  of a range is found by scanning it (lowest precedence at depth 0, through parentheses), as the AST
  would have it. A `?:` parses its condition as a value first; when the `?` shows up, that code is
  undone and the range is compiled again as a jump.
- **Collapsed tokens**: the tokenizer gives `a[1]` as one identifier; its index text is tokenized
  and appended after the program's tokens, then parsed as a range.

## Results

`python3 Tools/DirectCompiler/compare.py` (LibreOffice 24.2): the disassembly of `ASF_Direct` is
the one of the AST pipeline for all 158 programs of the corpus (8 written by hand, 150 generated with
random expressions, conditions, loops, indexes and properties). Compile times, tokenizer included,
milliseconds:

| Script | Tokenize | Tokens → AST → bytecode | Tokens → bytecode (direct) | Speed-up |
|---|---:|---:|---:|---:|
| 300 × `n = n + 1` | 137 | 1272 | 357 | 3.6× |
| 300 varied statements | 725 | 7204 | 2357 | 3.1× |

The tokenizer is now a third of the direct compiler's time; it is the next thing to speed up.

## Running it

- **LibreOffice**: `python3 Tools/DirectCompiler/compare.py` (parity, then timings); `-v` prints the
  code of any difference; `--export` writes `corpus.txt`, `bench1.vas` and `bench2.vas`.
- **Excel**: import the ASF classes (with `ASF_BYTECODE = 1`), `ASF_Direct.cls` and
  `DirectCompare.bas` into a workbook saved next to `corpus.txt`, `bench1.vas` and `bench2.vas`, and
  run `RunDirectComparison`; the results go to the Immediate window.
