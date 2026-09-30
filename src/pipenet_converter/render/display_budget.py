"""Fair display-only budgets; never use sampled geometry for graph extraction."""
from __future__ import annotations

from collections.abc import Hashable, Mapping


def bundle_quotas(counts: Mapping[Hashable, int], budget: int) -> dict[Hashable, int]:
    """Max-min allocation: retain small bundles whole regardless of DXF order."""
    if budget < 0 or any(n < 0 for n in counts.values()):
        raise ValueError('Display counts and budget must be nonnegative.')
    remaining = min(budget, sum(counts.values()))
    pending = sorted(counts, key=lambda key: (counts[key], repr(key)))
    result = {}
    for i, key in enumerate(pending):
        quota = min(counts[key], remaining // (len(pending)-i))
        result[key] = quota
        remaining -= quota
    return result


def display_index(index: int, total: int, quota: int) -> bool:
    """Evenly sample a large bundle, including both ends when quota permits."""
    if quota >= total:
        return True
    if quota <= 1:
        return quota == 1 and index == 0
    k = (index*(quota-1)+total-2)//(total-1)
    return k*(total-1)//(quota-1) == index
