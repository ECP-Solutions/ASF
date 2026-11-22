"""Tables with pandas for worksheet formulas: summaries, group totals, pivots, joins and queries.

A range with its header row arrives as a pandas DataFrame (the pandas.DataFrame hint); a DataFrame
back fills the cells with its column names. load_table keeps a table in Python and its cell shows a
handle such as <DataFrame#3>: the other functions take either that cell or a range."""
import pandas
from asfpy import xl_func

AGGREGATES = ('sum', 'mean', 'count', 'min', 'max', 'median')


def _frame(table):
    """A DataFrame from a handle's object or from the rows of a range (header first)."""
    if isinstance(table, pandas.DataFrame):
        return table
    return pandas.DataFrame(table[1:], columns=table[0])


@xl_func(name='PY.TABLE')
def load_table(table) -> object:
    """Keeps the table in Python: the cell shows a handle the other table functions take."""
    return _frame(table)


@xl_func(name='PY.DESCRIBE')
def table_summary(table):
    """Count, mean, spread and quartiles of each numeric column (pandas describe)."""
    return _frame(table).describe().round(4)


@xl_func(name='PY.GROUP')
def group_total(table, by: str, value: str, how: str = 'sum'):
    """One row per value of the column by, with the sum (or mean, count, min, max, median) of value."""
    if how not in AGGREGATES:
        raise ValueError('how is one of %s' % ', '.join(AGGREGATES))
    return _frame(table).groupby(by, as_index=False)[value].agg(how).sort_values(value, ascending=False)


@xl_func(name='PY.PIVOT')
def pivot(table, rows: str, columns: str, values: str, how: str = 'sum'):
    """A pivot table of values by rows and columns, with totals."""
    p = pandas.pivot_table(_frame(table), index=rows, columns=columns, values=values,
                           aggfunc=how, margins=True, margins_name='Total', fill_value=0)
    p.columns = [str(c) for c in p.columns]
    return p


@xl_func(name='PY.JOIN')
def join_tables(left, right, on: str, how: str = 'left'):
    """The rows of left with the columns of right whose on column matches (left, inner, outer)."""
    return _frame(left).merge(_frame(right), on=on, how=how)


@xl_func(name='PY.QUERY')
def query(table, condition: str):
    """The rows that meet a pandas condition, such as  Region == "North" and Qty > 10."""
    return _frame(table).query(condition)


@xl_func(name='PY.TOP')
def top_n(table, column: str, n: int = 5):
    """The n rows with the largest values of column."""
    return _frame(table).nlargest(n, column)


@xl_func(name='PY.CORR')
def correlation(table):
    """The correlation of every pair of numeric columns."""
    return _frame(table).select_dtypes('number').corr().round(4)
