"""The checks of the ASF add-in, shared by the LibreOffice checker (run_lo.py) and the test
workbook for Excel (make_workbook.py).

A case calls a worksheet function with arguments and expects a result:
  - an argument is a value, a Range of cell values (a list of rows), or an Error cell;
  - the result is a value, an Error, or a list of rows (an array formula over that many cells).
Cases with lo_only need the library folder of LIBRARY or a low step limit, which the Excel
workbook can't set up; the LibreOffice checker runs them all.
"""


class Range:
    """Cells passed as a range: a list of rows."""
    def __init__(self, rows):
        self.rows = rows


class Error:
    """An error value: in a cell, or as the result of a function."""
    CODES = {"#NULL!": 2000, "#DIV/0!": 2007, "#VALUE!": 2015, "#REF!": 2023, "#NAME?": 2029, "#NUM!": 2036, "#N/A": 2042}
    FORMULAS = {"#N/A": "=NA()", "#DIV/0!": "=1/0", "#VALUE!": '="a"+1'}

    def __init__(self, text):
        self.text = text

    def __eq__(self, other):
        return isinstance(other, Error) and other.text == self.text

    def __repr__(self):
        return self.text


# scripts of the library folder, for ASF_CALL
LIBRARY = {
    "double.vas": "return $1 * 2;",
    "util.vas": "fun twice(x) { return x * 2; };\nexport { twice };",
    "uses_util.vas": "import { twice } from './util.vas';\nreturn twice($1) + 1;",
    # scripts with a header, for the ribbon
    "report.vas": "// @asf name: A report\n// @asf args: selection\nreturn $1.length;",
    "bad.vas": "// @asf name: Bad\nthrow 'nope';",
    "nothing.vas": "let x = 1;",
    "typeof.vas": "return typeof $1;",
    "obj.vas": "// a comment before the header ends nothing\n// @asf name: Object\nreturn {n: 2};",
}

# the library as the ribbon lists it: header names, else file names, sorted without case
RIBBON_NAMES = "A report|Bad|double|nothing|Object|typeof|uses_util|util|"

# (script file, selection, the message of Run script)
RIBBON_RUNS = [
    ("report.vas", Range([[1], [2], [3]]), "Result: 3"),
    ("report.vas", None, "The script stopped: Cannot read properties of undefined (reading 'length')"),
    ("bad.vas", None, "The script stopped: nope"),
    ("nothing.vas", None, "Done."),
    ("obj.vas", None, "Result: { n: 2 }"),
    ("typeof.vas", Range([[5]]), "Result: undefined"),   # no args header: no $1
    ("no_such.vas", None, None),
]

# (source, key, the header value)
HEADERS = [
    ("// @asf name: X\n// @asf ARGS: Selection\nreturn 1;", "args", "Selection"),
    ("// @asf name: X\n// @asf ARGS: Selection\nreturn 1;", "appAccess", ""),
    ("\n// @asf appAccess:  true \r\nreturn 1;", "appaccess", "true"),
    ("return 1;\n// @asf name: late", "name", ""),
    ("//@asf name: no space", "name", "no space"),
]

CASES = [
    # values in, values out
    dict(name="arithmetic", fn="ASF_RUN", args=["return ($1 + $2) * 3 - $1 / 2;", 7, 5], expected=32.5),
    dict(name="whole numbers as doubles: % on them", fn="ASF_RUN", args=["return $1 % 4 + $2;", 7, 1], expected=4),
    dict(name="text", fn="ASF_RUN", args=["return $1 + ' ' + $2;", "a", "b"], expected="a b"),
    dict(name="boolean result", fn="ASF_RUN", args=["return $1 > $2;", 7, 5], expected=True),
    dict(name="no arguments", fn="ASF_RUN", args=["return 6 * 7;"], expected=42),
    dict(name="undefined gives an empty string", fn="ASF_RUN", args=["return $9;"], expected=""),
    # ranges in
    dict(name="one cell is its value", fn="ASF_RUN", args=["return typeof $1;", Range([[3]])], expected="number"),
    dict(name="an empty cell is undefined", fn="ASF_RUN", args=["return typeof $1;", Range([[None]])], expected="undefined"),
    dict(name="a range is an array of rows", fn="ASF_RUN", args=["return $1.length + 'x' + $1[1].length;", Range([[1, 2], [3, 4], [5, 6]])], expected="3x2"),
    dict(name="a cell of a range", fn="ASF_RUN", args=["return $1[2][1] * 10;", Range([[1, 2], [3, 4]])], expected=30),
    dict(name="a column is a flat array", fn="ASF_RUN", args=["return $1.reduce(fun(s, x) { return s + x; }, 0);", Range([[1], [2], [3]])], expected=6),
    dict(name="a column filtered and summed", fn="ASF_RUN", args=["return $1.filter(fun(x){ return x > 2 }).reduce(fun(a,x){ return a+x }, 0)", Range([[1], [2], [3], [4], [5]])], expected=12),
    dict(name="a row is a flat array", fn="ASF_RUN", args=["return $1.join('-');", Range([[1, 2, 3]])], expected="1-2-3"),
    dict(name="a column back as a column", fn="ASF_RUN", args=["return $1.map(fun(x) { return x * 10; });", Range([[1], [2]])], expected=[[10], [20]]),
    dict(name="an error cell as an argument is the result", fn="ASF_RUN", args=["return 1;", Range([[Error("#N/A")]])], expected=Error("#N/A")),
    dict(name="an error inside a range is its text", fn="ASF_RUN", args=["return $1[2];", Range([[1], [Error("#N/A")]])], expected="#N/A"),
    # arrays out
    dict(name="an array of rows is a grid", fn="ASF_RUN", args=["return [[1, 2], [3, 4]];"], expected=[[1, 2], [3, 4]]),
    dict(name="a flat array is a column", fn="ASF_RUN", args=["return [1, 2, 3];"], expected=[[1], [2], [3]]),
    dict(name="short rows are padded", fn="ASF_RUN", args=["return [[1], [2, 3]];"], expected=[[1, ""], [2, 3]]),
    dict(name="a range back as it came", fn="ASF_RUN", args=["return $1;", Range([["a", 1], ["b", 2]])], expected=[["a", 1], ["b", 2]]),
    dict(name="more cells than results: empty", fn="ASF_RUN", args=["return [1, 2];"], expected=[[1, ""], [2, ""], ["", ""]]),
    dict(name="fewer cells than results: cut", fn="ASF_RUN", args=["return [[1, 2, 3], [4, 5, 6]];"], expected=[[1, 2]]),
    dict(name="an empty array is an empty string", fn="ASF_RUN", args=["return [];"], expected=""),
    dict(name="an object is its text", fn="ASF_RUN", args=["return {a: 1, b: 'x'};"], expected="{ a: 1, b: 'x' }"),
    dict(name="a nested array in a cell is its text", fn="ASF_RUN", args=["return [[1, [2, 3]]];"], expected=[[1, "[ 2, 3 ]"]]),
    # errors
    dict(name="an error", fn="ASF_RUN", args=["let o = null; return o.x;"], expected=Error("#VALUE!")),
    dict(name="a throw no catch takes", fn="ASF_RUN", args=["throw 'bad';"], expected=Error("#VALUE!")),
    dict(name="a compile error", fn="ASF_RUN", args=["return (1 + ;"], expected=Error("#VALUE!")),
    dict(name="the next call after an error", fn="ASF_RUN", args=["return 'fine';"], expected="fine"),
    dict(name="an import outside the library", fn="ASF_RUN", args=["import { x } from '../outside.vas'; return 1;"], expected=Error("#VALUE!")),
    dict(name="a script name with a folder", fn="ASF_CALL", args=["../double", 1], expected=Error("#VALUE!")),
    dict(name="a script that is not there", fn="ASF_CALL", args=["no_such_script", 1], expected=Error("#VALUE!")),
    # regex functions
    dict(name="RX_TEST", fn="RX_TEST", args=["AB-1234", r"^[A-Z]{2}-\d{4}$"], expected=True),
    dict(name="RX_EXTRACT", fn="RX_EXTRACT", args=["12 kg", r"(\d+)\s*kg", 1], expected="12"),
    dict(name="RX_REPLACE", fn="RX_REPLACE", args=["a   b", r"\s+", " "], expected="a b"),
    dict(name="RX_COUNT", fn="RX_COUNT", args=["one two three", r"\b\w+\b"], expected=3),
    # the library and the step limit
    dict(name="a script of the library", fn="ASF_CALL", args=["double", 21], expected=42, lo_only=True),
    dict(name="with its extension", fn="ASF_CALL", args=["double.vas", 2], expected=4, lo_only=True),
    dict(name="a library script importing another", fn="ASF_CALL", args=["uses_util", 5], expected=11, lo_only=True),
    dict(name="ASF_RUN importing from the library", fn="ASF_RUN", args=["import { twice } from './util.vas'; return twice($1);", 4], expected=8, lo_only=True),
    dict(name="an endless loop stops", fn="ASF_RUN", args=["while (true) { }"], expected=Error("#VALUE!"), lo_only=True),
]


def shape(expected):
    """(rows, columns) of an expected list of rows; (1, 1) for one value."""
    if isinstance(expected, list):
        return len(expected), max(len(r) for r in expected)
    return 1, 1
