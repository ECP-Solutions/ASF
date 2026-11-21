"""The functions of the checks of phases 1a to 4. Copy this file into the Python folder
(%APPDATA%\\ASF\\python by default) before loading asfpy64.xll or asfpy32.xll."""
import collections
import datetime
import itertools
import math
import sys
import threading
import time

from asfpy import export, xl_app, xl_func, xl_macro

_volatile_calls = itertools.count(1)


@xl_func
def spike_version():
    """The Python that runs inside Excel."""
    return "Python %s, %d-bit" % (sys.version.split()[0], 64 if sys.maxsize > 2 ** 32 else 32)


@xl_func
def spike_add(a, b):
    """Adds two numbers."""
    return a + b


@xl_func
def spike_args3(a, b, c):
    """Shows three arguments as Python sees them."""
    return "%r | %r | %r" % (a, b, c)


@xl_func
def spike_sum7(a, b, c, d, e, f, g):
    """Adds seven numbers; empty cells count as 0."""
    return sum(x or 0 for x in (a, b, c, d, e, f, g))


@xl_func
def spike_shape(x):
    """What a range becomes in Python."""
    if isinstance(x, list):
        if x and isinstance(x[0], list):
            return "%d rows x %d columns" % (len(x), len(x[0]))
        return "list of %d" % len(x)
    return type(x).__name__


@xl_func
def spike_grid(rows, cols):
    """A grid of rows x columns: 11, 12... (an array formula)."""
    return [[(r + 1) * 10 + c + 1 for c in range(int(cols))] for r in range(int(rows))]


@xl_func
def spike_column(x):
    """A range back as one column of its values times 10."""
    flat = x if isinstance(x, list) else [x]
    return [[v * 10 if isinstance(v, float) else v] for v in (y for row in flat for y in (row if isinstance(row, list) else [row]))]


@xl_func
def spike_fail():
    """Raises an error on purpose: the cell shows #VALUE!."""
    raise ValueError("on purpose")


@xl_func
def spike_sqrt(x):
    """math.sqrt: #VALUE! for a negative number."""
    return math.sqrt(x)


@xl_func
def spike_text(s):
    """Text round trip, with characters beyond Latin-1."""
    return "%s · ü · 中文 · %d chars" % (s, len(s))


@xl_func
def spike_numpy(x):
    """The 90th percentile with numpy, when numpy is installed."""
    import numpy
    return float(numpy.percentile(numpy.asarray(x, dtype=float), 90))


@xl_func(volatile=True)
def spike_volatile():
    """Counts its calls: volatile, so each recalculation of the sheet counts once more."""
    return next(_volatile_calls)


@xl_func(thread_safe=True)
def spike_threadsafe(x):
    """x times 2, on Excel's calculation threads."""
    return x * 2


@xl_func
def spike_thread():
    """The thread a function that isn't thread-safe runs on: Excel's main thread."""
    return threading.current_thread().name


@xl_func(thread_safe=True, macro=True)
def spike_refused():
    """Never registered: Excel refuses thread-safe macro equivalents ($ with #)."""
    return "registered"


@xl_func
def spike_hints(n: int, when: datetime.date, text: str, values: list):
    """What type hints make of cell values."""
    return "%r | %r | %r | %r" % (n, when, text, values)


@xl_func
def spike_nextday(when: datetime.date):
    """The day after a date; a date comes back as Excel's serial number."""
    return when + datetime.timedelta(days=1)


@xl_func
def spike_counter(text):
    """A collections.Counter of the letters of text: not a cell value, so the cell shows a handle."""
    return collections.Counter(text)


@xl_func
def spike_most(counter, n: int):
    """The n-th most common letter of a Counter, given its handle."""
    return counter.most_common(n)[-1][0]


@xl_func
async def spike_async(x):
    """x times 2 after 50 ms, with asyncio: an asynchronous function."""
    import asyncio
    await asyncio.sleep(0.05)
    return x * 2


@xl_func(asynchronous=True)
def spike_async_thread(x):
    """x plus 1 after 50 ms, on a thread of the add-in."""
    time.sleep(0.05)
    return x + 1


@xl_func
def spike_matrix(a: 'numpy.ndarray'):
    """A range as a numpy array: the sum of each column, as a column (needs numpy)."""
    return a.sum(axis=0) if a.ndim == 2 else a.sum()


@xl_func
def spike_frame(table: 'pandas.DataFrame'):
    """A table with headers as a pandas DataFrame: its column names and row count (needs pandas)."""
    return '%s; %d rows' % (', '.join(table.columns), len(table))


@xl_func
def spike_describe(table: 'pandas.DataFrame'):
    """pandas' describe() of a table with headers, back to the cells with its index (needs pandas)."""
    return table.describe()


@xl_macro
def spike_macro(a, b):
    """A command: Application.Run "spike_macro", 2, 3 gives 5."""
    return a + b


@xl_macro
def spike_app():
    """A command that reaches Excel with xl_app() (needs pywin32): the name of the active sheet."""
    return xl_app().ActiveSheet.Name


@export
def add(a, b):
    return a + b


def Which():
    """Not marked: an ASF script reaches it as py.module('spike').Which(), where a lower-cased name
    would reach which() instead."""
    return "Which"


def which():
    return "which"


@export
def scale(values: list, factor: float):
    """Each value times factor: =PY_CALL("spike.scale", A1:A3, 10)."""
    return [v * factor for v in values]


@export
def squares(n):
    """A generator: VBA gets an object it can go through with For Each."""
    return (i * i for i in range(int(n)))


@export
def echo(*args):
    return list(args)


@export
def mean(values):
    flat = [v for row in values for v in (row if isinstance(row, list) else [row])]
    return sum(flat) / len(flat)
