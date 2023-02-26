from __future__ import annotations
from typing import NamedTuple


class ShapeShadowProps(NamedTuple):
    use: str  # Shadow
    blur: str  # ShadowBlur
    color: str  # ShadowColor
    transparence: str  # ShadowTransparence
    dist_x: str  # ShadowXDistance
    dist_y: str  # ShadowYDistance
