# Regex corpora

Randomly generated regular expression cases for `ASF_RegexEngine`, with the output they must produce. They check the engine against JavaScript's `RegExp` and against a previous version of itself, in Excel, in the LibreOffice edition or from a script.

| Corpus | Cases | Content |
|---|---|---|
| `cases/c1.txt` | 3,000 | literals, classes, groups, alternation, quantifiers, lookarounds, numbered backreferences, `\b` |
| `cases/c2.txt` | 6,000 | c1 plus code point escapes (`\x61`, `\u0062`), Unicode subjects, `^`/`$` anywhere, shorthand negations in classes |
| `cases/c3.txt` | 6,000 | c2 plus named groups and `\k<name>`, more alternatives per group |
| `cases/extra.txt` | 768 | 47 hand-picked patterns (the benchmark workloads, inline flags, atomic groups, named groups, empty patterns...) each with 16 subject and flag sets; some use ASF extensions that JavaScript rejects, so it has full output only |

Expected output:

- `expected/js/c1.js.txt` .. `c3.js.txt`: what JavaScript returns (Node.js 22, V8). ASF matches it on every case.
- `expected/full.zip`: `c1.full.txt` .. `c3.full.txt` and `extra.full.txt`, the output of every public method of the current engine (see Formats). Use it to check that a change to the engine changes no result.

## In Excel

1. Import the ASF class modules as usual (`src/*.cls`) and `test/regex-corpus/RegexCorpus.bas`.
2. In the Immediate window:

   ```vb
   ? RunCorpusSuite("C:\ASF\test\regex-corpus")
   ```

   It runs c1, c2 and c3, writes the output to `out\` in that folder, prints the differing cases of each corpus and returns `True` when none differ.

For the full check, unzip `expected\full.zip` (for example into `expected\full\`) and run, per corpus:

```vb
RunCorpusFull "C:\ASF\test\regex-corpus\cases\c2.txt", "C:\ASF\test\regex-corpus\out\c2.full.txt"
? CompareCorpusOutput("C:\ASF\test\regex-corpus\out\c2.full.txt", "C:\ASF\test\regex-corpus\expected\full\c2.full.txt")
```

`CompareCorpusOutput` prints the first differing lines (an optional third argument sets how many) and returns how many differ. Case 816 of c1 backtracks until the step limit on every call, so `RunCorpusFull` on c1 takes several minutes.

## In LibreOffice

From the repository root, with LibreOffice and `python3-uno` installed:

```sh
python3 test/regex-corpus/tools/run_lo.py              # c1 c2 c3 against JavaScript
python3 test/regex-corpus/tools/run_lo.py full c2 extra # every public method, against expected/full.zip
```

It builds the engine from `src/` (see `libreoffice/tools/build.py`), runs `RegexCorpus.bas` in a headless LibreOffice with a throw-away profile, writes the output to `test/regex-corpus/out/` and compares it, printing each differing case decoded. Options: `--repo=PATH` takes `src/` from another checkout, `--out=DIR` writes the output elsewhere, `--show=N` sets how many differences are printed. The exit status is 0 when every case matches. Corpora run one after another; to use several cores, start one command per corpus with different `--out` folders. In full mode, LibreOffice grows to about 5 GB of memory on c1 (case 816 backtracks to the step limit), so leave room for it when running corpora side by side. Times on one core of the machine that produced the expected output: about 2 to 3 minutes per corpus in JavaScript mode, 5 to 15 minutes in full mode.

`RegexCorpus.bas` creates engines with `New ASF_RegexEngine`, so in a LibreOffice document it has to be in the `ASF` library itself; `run_lo.py` does that for you.

## Comparing two versions of the engine

The expected full output belongs to one version of the engine. To compare any two versions, produce the full output of both and compare them:

```sh
python3 test/regex-corpus/tools/run_lo.py full c1 c2 c3 extra --repo=../asf-before --out=/tmp/before
python3 test/regex-corpus/tools/run_lo.py full c1 c2 c3 extra --out=/tmp/after
python3 test/regex-corpus/tools/compare.py test/regex-corpus/cases/c2.txt /tmp/before/c2.full.txt /tmp/after/c2.full.txt
```

In Excel, run `RunCorpusFull` in both workbooks and compare the two files with `CompareCorpusOutput`. When a change is meant to alter results, regenerate `expected/full.zip` from the new output once the differences are understood.

## Regenerating the cases and the JavaScript output

The generators are deterministic; these commands reproduce the files exactly:

```sh
cd test/regex-corpus
python3 tools/generators/gen_c1.py 7 3000 cases/c1.txt
python3 tools/generators/gen_c2.py 11 6000 cases/c2.txt
python3 tools/generators/gen_c3.py 21 6000 cases/c3.txt
node tools/js_run.js cases/c1.txt expected/js/c1.js.txt     # likewise c2, c3
```

Another seed or count gives a new corpus; produce its JavaScript output with `js_run.js` and run it like the others. A hand-written case is one line, encoded for example with:

```sh
python3 -c "import sys; h=lambda s: ''.join(f'{ord(c):04x}' for c in s); print(h(sys.argv[1]), h(sys.argv[2]), sys.argv[3], sep='\t')" '(a|ab)c' 'xabc' 'i'
```

## Formats

All text fields are UTF-16 code units in lowercase hex, four digits each, so any char survives the files.

- **Cases**: `pattern<TAB>subject<TAB>flags`, flags a subset of `ims` (ignoreCase, multiline, dotAll).
- **JavaScript mode** (`RunCorpusJs`, `js_run.js`), one line per case, tab-separated: `M` and the captures of the first match separated by commas (an unset group is empty), or `NOMATCH`; the full text of every match of a global search, comma-separated, or `NONE`; the subject with every match replaced by `<$&>`; the same with `<$&|$1>`. `ERR` when the pattern is rejected. `XERR` with the error when a call raises (step limit of 200,000).
- **Full mode** (`RunCorpusFull`), one line per case, three sections: as compiled, after flipping `ignoreCase` and `multiline` on the compiled engine (`|F`), and after `Init` of another pattern and back, served from the pattern cache (`|C`). Each section holds `E` `Exec`, `T` `test`, `A` `ExecAll`, `R` `Replace` with seven templates (`<$&>`, `<$&|$1>`, ``[$`|$']``, `$2$1$$`, `$<n>$01$10x$9$<`, `plain`, `$0$&$`), `R2`/`R3` two repeated replacements, `S` `Split` and `@` `ExecAt` at positions 1, 2, the length and the length + 1. Captures are `[c0,c1,...]`, `-` is `Nothing`, `{name=value;...}` the named captures after the call. `ERR` with the error when `Init` raises, `XERR` when a call does.
