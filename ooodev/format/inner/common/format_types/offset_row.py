from __future__ import annotations
from dataclasses import dataclass
from ooodev.utils.data_type.intensity import Intensity


@dataclass(unsafe_hash=True)
class OffsetRow(Intensity):
    """Represents a Row Offset value from ``0`` to ``100``."""

    pass
