from __future__ import annotations
from typing import TYPE_CHECKING
from ooodev.utils.data_type.generic_unit_size_pos import GenericUnitSizePos

# do not import from ooodev.unit or will cause circular import.
from ooodev.units.unit_px import UnitPX

if TYPE_CHECKING:
    from ooodev.units.unit_obj import UnitT


class SizePosPX(GenericUnitSizePos[UnitPX, float]):
    """
    Size and Position in ``px`` units.

    .. versionadded:: 0.39.0
    """

    def __init__(self, x: UnitPX, y: UnitPX, width: UnitPX, height: UnitPX) -> None:
        """
        Constructor

        Args:
            x (UnitPX): Width value as ``px``.
            y (UnitPX): Width value as ``px``.
            width (UnitPX): Width value as ``px``.
            height (UnitPX): Height value as ``px``.
        """
        super().__init__(x, y, width, height)

    @classmethod
    def from_unit_val(
        cls, x: UnitT | float | int, y: UnitT | float | int, width: UnitT | float | int, height: UnitT | float | int
    ) -> SizePosPX:
        """
        Get instance from ``UnitT`` or float value.

        Args:
            x (UnitT, float, int): ``UnitT`` or float value. If float then it is assumed to be in ``px`` units.
            y (UnitT, float, int): ``UnitT`` or float value. If float then it is assumed to be in ``px`` units.
            width (UnitT, float, int): ``UnitT`` or float value. If float then it is assumed to be in ``px`` units.
            height (UnitT, float, int): ``UnitT`` or float value. If float then it is assumed to be in ``px`` units.

        Returns:
            SizePosPX:
        """
        return cls(
            UnitPX.from_unit_val(x), UnitPX.from_unit_val(y), UnitPX.from_unit_val(width), UnitPX.from_unit_val(height)
        )
