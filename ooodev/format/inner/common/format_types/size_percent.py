from __future__ import annotations
from dataclasses import dataclass
from ooodev.utils.data_type.width_height_percent import WidthHeightPercent


@dataclass(frozen=True)
class SizePercent(WidthHeightPercent):
    """Size in percent values"""

    pass
