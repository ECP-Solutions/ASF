"""The functions of the phase 1a spike. Copy this file into the Python folder
(%APPDATA%\\ASF\\python by default) before loading asfpy64.xll or asfpy32.xll."""
import math
import sys

from asfpy import export, xl_func


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


@export
def add(a, b):
    return a + b


@export
def echo(*args):
    return list(args)


@export
def mean(values):
    flat = [v for row in values for v in (row if isinstance(row, list) else [row])]
    return sum(flat) / len(flat)
