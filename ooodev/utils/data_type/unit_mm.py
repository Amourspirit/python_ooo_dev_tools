from __future__ import annotations
from dataclasses import dataclass
from .base_float_value import BaseFloatValue
from ..unit_convert import UnitConvert


@dataclass(unsafe_hash=True)
class UnitMM(BaseFloatValue):
    """Unit in ``mm`` units."""

    def __post_init__(self):
        if not isinstance(self.value, float):
            object.__setattr__(self, "value", float(self.value))

    def get_value_mm100(self) -> int:
        """
        Gets instance value converted to Size in ``1/100th mm`` units.

        Returns:
            int: Value in ``1/100th mm`` units.
        """
        return UnitConvert.convert_mm_mm100(self.value)

    def from_mm100(value: int) -> UnitMM:
        """
        Get instance from ``1/100th mm`` value.

        Args:
            value (int): ``1/100th mm`` value.

        Returns:
            UnitMM:
        """
        return UnitMM(float(UnitConvert.convert_mm100_mm(value)))
