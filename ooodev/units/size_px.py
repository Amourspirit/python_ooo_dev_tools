from __future__ import annotations
from ooodev.utils.data_type.generic_unit_size import GenericUnitSize

# do not import from ooodev.unit or will cause circular import.
from ooodev.units.unit_px import UnitPX


class SizePX(GenericUnitSize[UnitPX, float]):
    """
    Size Width and Height in ``px`` units.

    .. versionadded:: 0.27.0
    """

    def __init__(self, width: UnitPX, height: UnitPX) -> None:
        """
        Constructor

        Args:
            width (UnitPX): Width value as ``px``.
            height (UnitPX): Height value as ``px``.
        """
        super().__init__(width, height)

    @classmethod
    def from_px(cls, width: float, height: float) -> SizePX:
        """
        Creates an instance from ``px`` values.

        Args:
            width (float): Width value as ``px``.
            height (float): Height value as ``px``.


        Returns:
            SizePX: An instance of SizePX.
        """
        return cls(UnitPX(width), UnitPX(height))
