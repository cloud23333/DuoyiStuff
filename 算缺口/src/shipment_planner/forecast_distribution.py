from __future__ import annotations

import math
from collections.abc import Sequence
from dataclasses import dataclass

from .forecast_level import (
    recent_mean,
    robust_level,
    weighted_mean,
    weighted_variance,
)

# var/mean at or below this ratio is treated as non-overdispersed → Poisson.
_POISSON_VAR_RATIO = 1.05
# Safety bound so a degenerate (mean, var) can never spin forever.
_MAX_QUANTILE_ITERS = 100_000
# Above this dispersion size the NegBin is indistinguishable from Poisson; fall
# back to avoid a slow near-Poisson CDF accumulation.
_MAX_NEGBIN_SIZE = 1.0e4


@dataclass(frozen=True, slots=True)
class DemandDistribution:
    """Predictive distribution of total demand over ``horizon`` days."""

    horizon: int
    mean: float
    variance: float

    def quantile(self, probability: float) -> float:
        return horizon_quantile(
            mean=self.mean, variance=self.variance, probability=probability
        )


def horizon_quantile(*, mean: float, variance: float, probability: float) -> float:
    """Smallest integer demand ``k`` with CDF(k) >= probability.

    Uses a Poisson inverse-CDF when demand is not overdispersed, otherwise a
    Negative-Binomial parametrised by (mean, variance). Pure stdlib.
    """
    clipped = min(1.0, max(0.0, probability))
    if mean <= 0.0:
        return 0.0
    if variance <= mean * _POISSON_VAR_RATIO:
        return _poisson_quantile(mean, clipped)
    return _negbin_quantile(mean, variance, clipped)


def _poisson_quantile(lam: float, probability: float) -> float:
    if lam <= 0.0:
        return 0.0
    log_lam = math.log(lam)
    cumulative = 0.0
    k = 0
    while k < _MAX_QUANTILE_ITERS:
        log_pmf = -lam + k * log_lam - math.lgamma(k + 1)
        cumulative += math.exp(log_pmf)
        if cumulative >= probability:
            return float(k)
        k += 1
    return float(k)


def _negbin_quantile(mean: float, variance: float, probability: float) -> float:
    success_prob = mean / variance
    if not 0.0 < success_prob < 1.0:
        return _poisson_quantile(mean, probability)
    size = (mean * mean) / (variance - mean)
    if size > _MAX_NEGBIN_SIZE:
        return _poisson_quantile(mean, probability)
    log_p = math.log(success_prob)
    log_q = math.log(1.0 - success_prob)
    lgamma_size = math.lgamma(size)
    cumulative = 0.0
    k = 0
    while k < _MAX_QUANTILE_ITERS:
        log_pmf = (
            math.lgamma(k + size)
            - lgamma_size
            - math.lgamma(k + 1)
            + size * log_p
            + k * log_q
        )
        cumulative += math.exp(log_pmf)
        if cumulative >= probability:
            return float(k)
        k += 1
    return float(k)


_SHORT_HALF_LIFE = 2.0
_LONG_HALF_LIFE = 5.0
_RECENT_DAYS = 5
# On a confirmed collapse, pin dispersion to the mean (Poisson): the floor at
# `level` and this cap collapse daily_var to exactly `level`, so the tail can't
# prop up the gap on a dying SKU.
_DROP_VAR_TO_MEAN_RATIO = 1.0


def _normalize(values: Sequence[float]) -> list[float]:
    return [max(0.0, float(value)) for value in values]


def _level(values: list[float], *, recent_drop: bool) -> float:
    """Spike-robust, recency-weighted daily level; capped on recent collapse."""
    level = robust_level(values)
    if recent_drop:
        level = min(level, recent_mean(values, days=_RECENT_DAYS))
    return level


def negbin_ewma_distribution(
    values: Sequence[float], *, horizon: int, recent_drop: bool = False
) -> DemandDistribution:
    base = _normalize(values)
    horizon = max(0, horizon)
    if not base or sum(base) <= 0 or horizon <= 0:
        return DemandDistribution(horizon=horizon, mean=0.0, variance=0.0)
    level = _level(base, recent_drop=recent_drop)
    # Dispersion from the RAW series so a one-off spike widens the interval.
    raw_mean = weighted_mean(base, _LONG_HALF_LIFE)
    daily_var = max(weighted_variance(base, _LONG_HALF_LIFE, mean=raw_mean), level)
    if recent_drop:
        daily_var = min(daily_var, level * _DROP_VAR_TO_MEAN_RATIO)
    return DemandDistribution(
        horizon=horizon, mean=level * horizon, variance=daily_var * horizon
    )


def poisson_ewma_distribution(
    values: Sequence[float], *, horizon: int, recent_drop: bool = False
) -> DemandDistribution:
    base = _normalize(values)
    horizon = max(0, horizon)
    if not base or sum(base) <= 0 or horizon <= 0:
        return DemandDistribution(horizon=horizon, mean=0.0, variance=0.0)
    level = _level(base, recent_drop=recent_drop)
    mean_h = level * horizon
    return DemandDistribution(horizon=horizon, mean=mean_h, variance=mean_h)


def hurdle_distribution(
    values: Sequence[float], *, horizon: int, recent_drop: bool = False
) -> DemandDistribution:
    base = _normalize(values)
    horizon = max(0, horizon)
    if not base or sum(base) <= 0 or horizon <= 0:
        return DemandDistribution(horizon=horizon, mean=0.0, variance=0.0)
    indicators = [1.0 if value > 0 else 0.0 for value in base]
    occurrence = weighted_mean(indicators, _LONG_HALF_LIFE)
    positives = [value for value in base if value > 0]
    if not positives:
        return DemandDistribution(horizon=horizon, mean=0.0, variance=0.0)
    level_size = robust_level(positives)
    raw_size = weighted_mean(positives, _LONG_HALF_LIFE)
    raw_size_var = weighted_variance(positives, _LONG_HALF_LIFE, mean=raw_size)
    if recent_drop:
        occurrence = min(occurrence, recent_mean(indicators, days=_RECENT_DAYS))
        level_size = min(level_size, recent_mean(base, days=_RECENT_DAYS) or level_size)
    day_mean = occurrence * level_size
    day_var = (
        occurrence * raw_size_var
        + occurrence * (1.0 - occurrence) * raw_size * raw_size
    )
    mean_h = day_mean * horizon
    var_h = max(day_var * horizon, mean_h)
    if recent_drop:
        var_h = mean_h
    return DemandDistribution(horizon=horizon, mean=mean_h, variance=var_h)
