from __future__ import annotations
from typing import NamedTuple


class ImageFlipProps(NamedTuple):
    vert: str  # VertMirrored
    hori_even: str  # HoriMirroredOnEvenPages
    hori_odd: str  #  HoriMirroredOnOddPages
