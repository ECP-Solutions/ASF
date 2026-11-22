"""Statistics with numpy for worksheet formulas: standard scores, outliers, trends and forecasts.

A range arrives as a numpy array (the numpy.ndarray hint), empty cells as nan, which the functions
skip. An array back fills the cells of the formula, one column for a list of values."""
import types

import numpy
from asfpy import export, xl_func


def _clean(values):
    v = numpy.asarray(values, dtype=float).ravel()
    return v[~numpy.isnan(v)]


@xl_func(name='PY.ZSCORES')
def zscores(values: 'numpy.ndarray'):
    """How many standard deviations each value lies from the mean (empty cells stay empty).

    Args:
        values: the numbers, a row or a column
    """
    v = numpy.asarray(values, dtype=float).ravel()
    clean = v[~numpy.isnan(v)]
    z = (v - clean.mean()) / clean.std(ddof=1)
    return [[None if numpy.isnan(x) else round(float(x), 6)] for x in z]


@xl_func(name='PY.OUTLIERS')
def outlier_flags(values: 'numpy.ndarray', k: float = 1.5):
    """TRUE for each value outside Tukey's fences: below Q1 - k*IQR or above Q3 + k*IQR.

    Args:
        values: the numbers, a row or a column
        k: how many interquartile ranges beyond the quartiles; 1.5 by default, 3 for far outliers
    """
    v = numpy.asarray(values, dtype=float).ravel()
    q1, q3 = numpy.percentile(_clean(v), [25, 75])
    low, high = q1 - k * (q3 - q1), q3 + k * (q3 - q1)
    return [[None if numpy.isnan(x) else bool(x < low or x > high)] for x in v]


@xl_func(name='PY.TREND')
def trend(y: 'numpy.ndarray', x: 'numpy.ndarray' = None):
    """The least-squares line through the values: slope, intercept and R², with headers.

    Args:
        y: the values
        x: their x values; 1, 2, 3... when left out
    """
    y = numpy.asarray(y, dtype=float).ravel()
    x = numpy.arange(1, len(y) + 1, dtype=float) if x is None else numpy.asarray(x, dtype=float).ravel()
    keep = ~(numpy.isnan(x) | numpy.isnan(y))
    slope, intercept = numpy.polyfit(x[keep], y[keep], 1)
    fitted = slope * x[keep] + intercept
    r2 = 1 - ((y[keep] - fitted) ** 2).sum() / ((y[keep] - y[keep].mean()) ** 2).sum()
    return [['slope', 'intercept', 'R2'], [float(slope), float(intercept), float(r2)]]


@xl_func(name='PY.FORECAST')
def forecast(y: 'numpy.ndarray', periods: int = 3, season: int = 0):
    """The next values of a series: its linear trend, plus the average seasonal pattern when season
    gives the length of a cycle (12 for months).

    Args:
        y: the series, oldest first
        periods: how many values to forecast; 3 by default
        season: the length of a cycle, such as 12 for months; 0 for none
    """
    y = _clean(y)
    t = numpy.arange(len(y), dtype=float)
    slope, intercept = numpy.polyfit(t, y, 1)
    future = numpy.arange(len(y), len(y) + periods, dtype=float)
    out = slope * future + intercept
    if season and len(y) >= 2 * season:
        residual = y - (slope * t + intercept)
        pattern = numpy.array([residual[i::season].mean() for i in range(season)])
        out = out + pattern[future.astype(int) % season]
    return [[round(float(v), 4)] for v in out]


@xl_func(name='PY.MOVAVG')
def moving_average(values: 'numpy.ndarray', window: int = 3):
    """The mean of each value and the window - 1 before it; empty for the first ones.

    Args:
        values: the series
        window: how many values each mean takes; 3 by default
    """
    v = numpy.asarray(values, dtype=float).ravel()
    out = numpy.convolve(v, numpy.ones(window) / window, mode='valid')
    return [[None]] * (window - 1) + [[round(float(x), 6)] for x in out]


@xl_func(name='PY.HISTOGRAM')
def histogram(values: 'numpy.ndarray', bins: int = 5):
    """How many values fall in each of bins equal intervals: from, to, count.

    Args:
        values: the numbers
        bins: how many intervals; 5 by default
    """
    counts, edges = numpy.histogram(_clean(values), bins=bins)
    return [['from', 'to', 'count']] + [[float(edges[i]), float(edges[i + 1]), int(c)] for i, c in enumerate(counts)]


@export
def summary(values):
    """For ASF scripts (py.call): an object with the count, mean, std, min and max of a list."""
    v = _clean(numpy.array([numpy.nan if x is None or x == '' else x for x in values], dtype=float))
    return types.SimpleNamespace(count=int(v.size), mean=float(v.mean()), std=float(v.std(ddof=1)),
                                 min=float(v.min()), max=float(v.max()))


@export
def above_mean(values):
    """For PY_CALL: how many values lie above their mean."""
    v = _clean(numpy.array([numpy.nan if x is None or x == '' else x for x in values], dtype=float))
    return int((v > v.mean()).sum())
