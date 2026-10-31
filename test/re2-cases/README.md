# RE2 test cases

[RE2](https://github.com/google/re2)'s exhaustive tests, as extracted by Sihlfall into [sihlfall/regex-test-cases](https://github.com/sihlfall/regex-test-cases) (BSD-3-Clause, from RE2 commit `03da4fc0`), run against `ASF_RegexEngine`. The 27 suites hold 5,716,884 cases (pattern, subject) built from small alphabets: literals, repetition with and without captures, empty strings, punctuation, line ends, character classes and UTF-8 subjects. Each case has RE2's leftmost-first result, unanchored and anchored at both ends, which is what a backtracking engine such as ASF computes; the leftmost-longest results of the source files are not used.

The cases are not copied here (RE2's results take 481 MB): `tools/prepare.py` clones the repository at the pinned commit `0ab3381` and converts it into `work/`.

## Translation

RE2 runs with its default flags. Where RE2 spells something differently from ASF, the pattern is translated to the ASF pattern with the same meaning; `compare.py` shows both.

| RE2 | ASF | Why |
|---|---|---|
| `.` | `[^\n]` | RE2's dot excludes only `\n`; ASF's also excludes `\r` |
| `\C` | `[\s\S]` | any byte in RE2; compared on ASCII subjects only (see below) |
| `\A`, `\z` | `^`, `$` | start and end of text (no multiline) |
| `\a` | `\x07` | bell |
| `\s`, `\S` | `[\t\n\f\r ]`, `[^\t\n\f\r ]` | RE2's `\s` has no `\v`; ASF's has |
| `[[:alpha:]]` and the other POSIX classes | ASCII ranges | not in ASF's syntax |
| `[]a]`, `[^]a]` | `[\]a]`, `[^\]a]` | a leading `]` is a member in RE2 |

RE2 reports byte offsets in UTF-8; `prepare.py` turns them into UTF-16 code unit offsets, and each group into its text, since ASF's public API returns group text. A group that did not take part and an empty group are both `""`.

## Checks and categories

Every case is checked twice: the leftmost match (the first `ExecAt` position that matches, as `Exec` finds it) and the match of `^(?:pattern)$` at position 1 (RE2's anchor at both ends). Each check compares the span and the text of every group. A check that differs falls into the first category that explains it:

| Category | Meaning |
|---|---|
| `utf16` | the subject has a char above U+FFFF: two UTF-16 code units for ASF (as for JavaScript without the `u` flag), one code point for RE2 |
| `bytes` | the pattern has `\C` and the subject is not ASCII: one byte in RE2, one char in ASF |
| `ecmascript` | ASF gives what JavaScript's `RegExp` gives (checked with Node.js): ASF follows ECMAScript where it differs from Perl and RE2, as in loops whose iterations can match empty |
| `rejected` | ASF rejected the pattern |
| `error` | ASF raised while matching |
| `groups` | same span, different group text |
| `span` | different match, or a match against no match |

The first three are known differences of semantics; the last four would be defects. Checks where RE2 matched part of a UTF-8 sequence (`\C` again) have no UTF-16 equivalent and are skipped.

## Running

Prepare the cases once (needs git and Python 3; about 20 s):

```sh
python3 test/re2-cases/tools/prepare.py
```

In LibreOffice, with `python3-uno` (and Node.js for the `ecmascript` category):

```sh
python3 test/re2-cases/tools/run_lo.py              # all 27 suites
python3 test/re2-cases/tools/run_lo.py 016 018      # some suites
```

It builds the engine from `src/`, runs `Re2Cases.bas` in a headless LibreOffice, writes `out/NNN.txt` and prints, per suite, the number of checks in each category and the first differences of each kind (`--show=N`). The whole set takes about 3 hours on one core of LibreOffice Basic; suites run one after another, so start one command per group of suites to use several cores. The exit status is 0 when every difference is in the `utf16`, `bytes` or `ecmascript` category.

In Excel, import the ASF class modules and `test/re2-cases/Re2Cases.bas`, run `RunRe2Suite "C:\ASF\test\re2-cases"` (or `RunRe2Cases` for one suite) and compare the output:

```sh
python3 test/re2-cases/tools/compare.py test/re2-cases/work test/re2-cases/out
```

`compare.py` also writes `out/summary.json` with the counts and examples.

## Files

- `Re2Cases.bas`: runs a prepared suite (`RunRe2Cases`) or all of them (`RunRe2Suite`).
- `tools/prepare.py`: clone, translation and conversion of RE2's results.
- `tools/run_lo.py`: LibreOffice runner. `tools/compare.py`: comparison and categories. `tools/js_check.js`: JavaScript results for the differing checks.
