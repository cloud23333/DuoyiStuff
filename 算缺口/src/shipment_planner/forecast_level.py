# src/shipment_planner/forecast_level.py
from __future__ import annotations

import statistics
from collections.abc import Sequence

ISOLATED_SPIKE_MULTIPLIER = 3.0
ISOLATED_SPIKE_MIN_QTY = 10.0
ISOLATED_SPIKE_CONTEXT_DAYS = 5
RECENT_DROP_DAYS = 5
RECENT_DROP_RATIO = 0.45
INTERMITTENT_ZERO_RATIO = 0.5
SUSTAINED_RISE_SHORT_HALF_LIFE = 2.0
SUSTAINED_RISE_LONG_HALF_LIFE = 5.0
SUSTAINED_RISE_RATIO = 1.5


def _alpha_from_half_life(half_life: float) -> float:
    if half_life <= 0:
        return 1.0
    return 1.0 - 0.5 ** (1.0 / half_life)


def _exp_weights(count: int, half_life: float) -> list[float]:
    """Normalised geometric weights, newest sample heaviest."""
    decay = 1.0 - _alpha_from_half_life(half_life)
    raw = [decay ** (count - 1 - i) for i in range(count)]
    total = sum(raw)
    if total <= 0:
        return [1.0 / count] * count
    return [value / total for value in raw]


def weighted_mean(values: Sequence[float], half_life: float) -> float:
    if not values:
        return 0.0
    weights = _exp_weights(len(values), half_life)
    return sum(w * float(v) for w, v in zip(weights, values))


def weighted_variance(
    values: Sequence[float], half_life: float, *, mean: float | None = None
) -> float:
    if len(values) < 2:
        return 0.0
    weights = _exp_weights(len(values), half_life)
    center = mean if mean is not None else weighted_mean(values, half_life)
    return sum(w * (float(v) - center) ** 2 for w, v in zip(weights, values))


ROBUST_LEVEL_HALF_LIFE = 5.0
WINSOR_UPPER_QUANTILE = 0.9
TRIM_FRACTION = 0.2


def winsorize(values: Sequence[float], upper_quantile: float = WINSOR_UPPER_QUANTILE) -> list[float]:
    """Cap values at the ``upper_quantile`` element (nearest-rank, inclusive at 1.0).

    The nearest-rank index floors, so on very short series the cap lands lower than
    the quantile suggests; this only biases the estimate downward, which is the
    intended conservative direction for spike-robust level estimation.
    """
    if not values:
        return []
    ordered = sorted(values)
    cap_index = min(len(ordered) - 1, int(upper_quantile * (len(ordered) - 1)))
    cap = ordered[cap_index]
    return [min(float(value), cap) for value in values]


def trimmed_mean(values: Sequence[float], trim_fraction: float = TRIM_FRACTION) -> float:
    if not values:
        return 0.0
    ordered = sorted(float(value) for value in values)
    cut = int(trim_fraction * len(ordered))
    core = ordered[cut : len(ordered) - cut] or ordered
    return statistics.fmean(core)


def robust_level(values: Sequence[float], half_life: float = ROBUST_LEVEL_HALF_LIFE) -> float:
    """Spike-robust daily level.

    Lower of two views so a one-off jump can enter neither:
      - a recency-weighted mean of the winsorized series (follows real trends), and
      - a robust central value, max(median, trimmed mean), acting as a cap.
    """
    base = [max(0.0, float(value)) for value in values]
    if not base:
        return 0.0
    follow = weighted_mean(winsorize(base), half_life)
    central_cap = max(statistics.median(base), trimmed_mean(base))
    return min(follow, central_cap)


def ewma(values: Sequence[float], half_life: float) -> float:
    if not values:
        return 0.0
    if half_life <= 0:
        return float(values[-1])
    alpha = _alpha_from_half_life(half_life)
    result = float(values[0])
    for value in values[1:]:
        result = alpha * value + (1.0 - alpha) * result
    return result


def recent_mean(values: Sequence[float], *, days: int) -> float:
    if days <= 0 or not values:
        return 0.0
    recent = values[-days:]
    return statistics.fmean(recent) if recent else 0.0


def clean_isolated_spikes(values: Sequence[float]) -> tuple[list[float], bool]:
    """Downweight one-off sales spikes not sustained by neighboring days."""
    cleaned = [float(value) for value in values]
    if len(cleaned) < ISOLATED_SPIKE_CONTEXT_DAYS:
        return cleaned, False

    changed = False
    radius = ISOLATED_SPIKE_CONTEXT_DAYS // 2
    for index in range(len(cleaned)):
        peer_values = cleaned[:index] + cleaned[index + 1 :]
        peer_positive = [value for value in peer_values if value > 0]
        peer_baseline = (
            float(statistics.median(peer_positive))
            if peer_positive
            else float(statistics.median(peer_values))
        )
        baseline = max(1.0, peer_baseline)
        local_peer_values = (
            cleaned[index - radius : index] + cleaned[index + 1 : index + radius + 1]
        )
        if cleaned[index] < max(3.0, baseline * ISOLATED_SPIKE_MULTIPLIER):
            continue
        if cleaned[index] < ISOLATED_SPIKE_MIN_QTY:
            continue
        local_support_threshold = min(
            cleaned[index] * 0.5, baseline * ISOLATED_SPIKE_MULTIPLIER
        )
        if any(value >= local_support_threshold for value in local_peer_values):
            continue
        cleaned[index] = baseline
        changed = True
    return cleaned, changed


def has_recent_drop(values: Sequence[float], *, stock_in_warehouse: float) -> bool:
    if stock_in_warehouse <= 0 or len(values) < 6:
        return False
    zero_ratio = sum(1 for value in values if value <= 0) / len(values)
    if zero_ratio >= INTERMITTENT_ZERO_RATIO:
        return False
    recent = values[-RECENT_DROP_DAYS:]
    previous = values[:-RECENT_DROP_DAYS]
    previous_positive = [value for value in previous if value > 0]
    if not previous_positive:
        return False
    baseline = float(statistics.median(previous_positive))
    if baseline < 2.0:
        return False
    return statistics.fmean(recent) <= baseline * RECENT_DROP_RATIO


def has_sustained_rise(values: Sequence[float]) -> bool:
    if len(values) < 4:
        return False
    cleaned, isolated = clean_isolated_spikes(values)
    if isolated:
        return False
    median_daily = float(statistics.median(cleaned)) if cleaned else 0.0
    baseline = max(1.0, median_daily)
    short = ewma(cleaned, SUSTAINED_RISE_SHORT_HALF_LIFE)
    long = ewma(cleaned, SUSTAINED_RISE_LONG_HALF_LIFE)
    return short > baseline * SUSTAINED_RISE_RATIO and long >= baseline
