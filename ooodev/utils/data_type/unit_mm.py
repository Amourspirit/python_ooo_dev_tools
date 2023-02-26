from __future__ import annotations
from typing import TypeVar, Type
from dataclasses import dataclass
from .base_float_value import BaseFloatValue
from ..unit_convert import UnitConvert, Length

_TUnitMM = TypeVar(name="_TUnitMM", bound="UnitMM")


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

    def get_value_pt(self) -> int:
        """
        Gets instance value converted to Size in `pt`` (points) units.

        Returns:
            int: Value in ``pt`` units.
        """
        return round(UnitConvert.convert(num=self.value, frm=Length.MM, to=Length.PT))

    @classmethod
    def from_mm100(cls: Type[_TUnitMM], value: int) -> _TUnitMM:
        """
        Get instance from ``1/100th mm`` value.

        Args:
            value (int): ``1/100th mm`` value.

        Returns:
            UnitMM:
        """
        inst = super(UnitMM, cls).__new__(cls)
        return inst.__init__(value)

    @classmethod
    def from_pt(cls: Type[_TUnitMM], value: int) -> _TUnitMM:
        """
        Get instance from ``pt`` (points) value.

        Args:
            value (int): ``pt`` value.

        Returns:
            UnitMM:
        """
        inst = super(UnitMM, cls).__new__(cls)
        return inst.__init__(float(UnitConvert.convert(num=value, frm=Length.PT, to=Length.MM)))
