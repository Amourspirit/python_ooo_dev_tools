from __future__ import annotations
from ooodev.utils.data_type.generic_unit_size import GenericUnitSize

# do not import from ooodev.unit or will cause circular import.
from ooodev.units.unit_mm100 import UnitMM100


class SizeMM100(GenericUnitSize[UnitMM100, int]):
    """
    Size Width and Height in ``1/100th mm`` units.

    .. versionadded:: 0.27.0
    """

    def __init__(self, width: UnitMM100, height: UnitMM100) -> None:
        """
        Constructor

        Args:
            width (UnitMM100): Width value as ``1/100th mm``.
            height (UnitMM100): Height value as ``1/100th mm``.
        """
        super().__init__(width, height)

    @classmethod
    def from_mm100(cls, width: int, height: int) -> SizeMM100:
        """
        Creates an instance from ``1/100th mm`` values.

        Args:
            width (int): Width value as ``1/100th mm``.
            height (int): Height value as ``1/100th mm``.

        Returns:
            SizeMM100: An instance of SizeMM100.
        """
        return cls(UnitMM100(width), UnitMM100(height))
