"""The functions of the checks of phases 1a and 1b. Copy this file into the Python folder
(%APPDATA%\\ASF\\python by default) before loading asfpy64.xll or asfpy32.xll."""
import itertools
import math
import sys
import threading

from asfpy import export, xl_func

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
    """The 90th percentile with numpy, when numpy is installed in the venv."""
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


@export
def add(a, b):
    return a + b


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
