from __future__ import annotations
from typing import Tuple, TYPE_CHECKING
from ..format.kind.format_kind import FormatKind as FormatKind

try:
    from typing import Protocol
except ImportError:
    from typing_extensions import Protocol


class UnitObj(Protocol):
    """Protolcol Class for units"""

    def get_value_mm(self) -> float:
        """
        Gets instance value converted to Size in ``mm`` units.

        Returns:
            int: Value in ``mm`` units.
        """
        ...

    def get_value_mm100(self) -> int:
        """
        Gets instance value converted to Size in ``1/100th mm`` units.

        Returns:
            int: Value in ``1/100th mm`` units.
        """
        ...

    def get_value_pt(self) -> float:
        """
        Gets instance value converted to Size in ``pt`` (point) units.

        Returns:
            int: Value in ``pt`` units.
        """
        ...

    def get_value_px(self) -> float:
        """
        Gets instance value in ``px`` (pixel) units.

        Returns:
            int: Value in ``px`` units.
        """
        ...
