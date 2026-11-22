"""Finance with numpy: a loan's schedule, compound growth, and a Monte Carlo simulation that runs
while Excel goes on calculating."""
import numpy
from asfpy import xl_func


@xl_func(name='PY.LOAN')
def loan_schedule(principal: float, annual_rate: float, years: int, per_year: int = 12):
    """Every payment of a fixed-rate loan: period, payment, interest, principal and balance."""
    n, r = years * per_year, annual_rate / per_year
    payment = principal / n if r == 0 else principal * r / (1 - (1 + r) ** -n)
    rows, balance = [['period', 'payment', 'interest', 'principal', 'balance']], principal
    for k in range(1, n + 1):
        interest = balance * r
        paid = payment - interest
        balance = max(balance - paid, 0.0)
        rows.append([k, round(payment, 2), round(interest, 2), round(paid, 2), round(balance, 2)])
    return rows


@xl_func(name='PY.CAGR')
def cagr(first: float, last: float, years: float):
    """The compound annual growth rate that turns first into last in years."""
    return (last / first) ** (1 / years) - 1


@xl_func(name='PY.MONTECARLO', asynchronous=True)
def monte_carlo(start: float, mean_return: float, volatility: float, years: int,
                runs: int = 10000, seed: int = 1):
    """The value of an investment after years of random returns (normal, yearly), as the 5th, 50th
    and 95th percentiles of runs simulations. Asynchronous: Excel goes on while it runs."""
    rng = numpy.random.default_rng(seed)
    returns = rng.normal(mean_return, volatility, size=(runs, years))
    final = start * numpy.prod(1 + returns, axis=1)
    p5, p50, p95 = numpy.percentile(final, [5, 50, 95])
    return [['P5', 'P50', 'P95'], [round(float(p5), 2), round(float(p50), 2), round(float(p95), 2)]]
