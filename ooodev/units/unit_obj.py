from __future__ import annotations
from typing import TYPE_CHECKING
from ..format.inner.kind.format_kind import FormatKind as FormatKind

if TYPE_CHECKING:
    try:
        from typing import Protocol
    except ImportError:
        from typing_extensions import Protocol
else:
    Protocol = object


class UnitObj(Protocol):
    """
    Protocol Class for units.

    .. _proto_unit_obj:

    UnitObj
    =======

    """

    def get_value_mm(self) -> float:
        """
        Gets instance value converted to Size in ``mm`` units.

        Returns:
            float: Value in ``mm`` units.
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
            float: Value in ``pt`` units.
        """
        ...

    def get_value_px(self) -> float:
        """
        Gets instance value in ``px`` (pixel) units.

        Returns:
            float: Value in ``px`` units.
        """
        ...
