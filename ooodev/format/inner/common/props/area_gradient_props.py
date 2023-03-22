from __future__ import annotations
from typing import NamedTuple


class AreaGradientProps(NamedTuple):
    style: str  # FillStyle such as FillStyle.GRADIENT
    step_count: str  # int, FillGradientStepCount
    name: str  # FillGradientName
    grad_prop_name: str  # str FillGradient
