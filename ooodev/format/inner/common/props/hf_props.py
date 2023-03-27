from __future__ import annotations
from typing import NamedTuple


class HfProps(NamedTuple):
    """Internal Properties"""

    on: str  # HeaderIsOn (bool)
    shared: str  # HeaderIsShared (bool, Same contents left and right)
    shared_first: str  # FirstIsShared (bool, same content on first page)
    margin_left: str  # HeaderLeftMargin (1/100th mm)
    margin_right: str  # HeaderRightMargin (1/100th mm)
    spacing: str  # HeaderBodyDistance (in 1/100th mm)
    spacing_dyn: str  # HeaderDynamicSpacing ( bool )
    height: str  # HeaderHeight (1/100th mm + 600)
    height_auto: str  # HeaderIsDynamicHeight ( bool )
