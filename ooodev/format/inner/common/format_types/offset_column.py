from __future__ import annotations
from dataclasses import dataclass
from ooodev.utils.data_type.intensity import Intensity


@dataclass(unsafe_hash=True)
class OffsetColumn(Intensity):
    """Represents a Column Offset value from ``0`` to ``100``."""

    pass
