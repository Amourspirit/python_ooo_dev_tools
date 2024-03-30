from __future__ import annotations
from typing import TYPE_CHECKING
from ooodev.utils.data_type.generic_unit_size_pos import GenericUnitSizePos

# do not import from ooodev.unit or will cause circular import.
from ooodev.units.unit_mm100 import UnitMM100

if TYPE_CHECKING:
    from ooodev.units.unit_obj import UnitT


class SizePosMM100(GenericUnitSizePos[UnitMM100, int]):
    """
    Size and Position in ``1/100 mm`` units.

    .. versionadded:: 0.39.0
    """

    def __init__(self, x: UnitMM100, y: UnitMM100, width: UnitMM100, height: UnitMM100) -> None:
        """
        Constructor

        Args:
            x (UnitMM100): Width value as ``1/100 mm``.
            y (UnitMM100): Width value as ``1/100 mm``.
            width (UnitMM100): Width value as ``1/100 mm``.
            height (UnitMM100): Height value as ``1/100 mm``.
        """
        super().__init__(x, y, width, height)

    @classmethod
    def from_unit_val(
        cls, x: UnitT | float | int, y: UnitT | float | int, width: UnitT | float | int, height: UnitT | float | int
    ) -> SizePosMM100:
        """
        Get instance from ``UnitT`` or float value.

        Args:
            x (UnitT, float, int): ``UnitT`` or float value. If float then it is assumed to be in ``1/100 mm`` units.
            y (UnitT, float, int): ``UnitT`` or float value. If float then it is assumed to be in ``1/100 mm`` units.
            width (UnitT, float, int): ``UnitT`` or float value. If float then it is assumed to be in ``1/100 mm`` units.
            height (UnitT, float, int): ``UnitT`` or float value. If float then it is assumed to be in ``1/100 mm`` units.

        Returns:
            SizePosMM100:
        """
        return cls(
            UnitMM100.from_unit_val(x),
            UnitMM100.from_unit_val(y),
            UnitMM100.from_unit_val(width),
            UnitMM100.from_unit_val(height),
        )
