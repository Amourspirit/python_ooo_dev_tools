from __future__ import annotations
from typing import TypeVar, Type
from dataclasses import dataclass
from .base_float_value import BaseFloatValue
from ..unit_convert import UnitConvert, Length

_TUnitInch = TypeVar(name="_TUnitInch", bound="UnitInch")


@dataclass(unsafe_hash=True)
class UnitInch(BaseFloatValue):
    """
    Unit in ``inch`` units.

    Supports ``UnitObj`` protcol.

    See Also:
        :ref:`proto_unit_obj`
    """

    def __post_init__(self):
        if not isinstance(self.value, float):
            object.__setattr__(self, "value", float(self.value))

    def get_value_pt(self) -> float:
        """
        Gets instance value converted to Size in ``pt`` (points) units.

        Returns:
            int: Value in ``pt`` units.
        """
        return UnitConvert.convert(num=self.value, frm=Length.IN, to=Length.PT)

    def get_value_mm(self) -> float:
        """
        Gets instance value converted to Size in ``mm`` units.

        Returns:
            int: Value in ``mm`` units.
        """
        return UnitConvert.convert(num=self.value, frm=Length.IN, to=Length.MM)

    def get_value_mm100(self) -> int:
        """
        Gets instance value converted to Size in ``1/100th mm`` units.

        Returns:
            int: Value in ``1/100th mm`` units.
        """
        return round(UnitConvert.convert(num=self.value, frm=Length.IN, to=Length.MM100))

    @classmethod
    def from_mm(cls: Type[_TUnitInch], value: float) -> _TUnitInch:
        """
        Get instance from ``mm`` value.

        Args:
            value (float): ``mm`` value.

        Returns:
            UnitInch:
        """
        inst = super(UnitInch, cls).__new__(cls)
        return inst.__init__(UnitConvert.convert(num=value, frm=Length.MM, to=Length.IN))

    @classmethod
    def from_mm100(cls: Type[_TUnitInch], value: int) -> _TUnitInch:
        """
        Get instance from ``1/100th mm`` value.

        Args:
            value (int): ``1/100th mm`` value.

        Returns:
            UnitInch:
        """
        inst = super(UnitInch, cls).__new__(cls)
        return inst.__init__(UnitConvert.convert(num=value, frm=Length.MM100, to=Length.IN))

    @classmethod
    def from_pt(cls: Type[_TUnitInch], value: float) -> _TUnitInch:
        """
        Get instance from ``pt`` (points) value.

        Args:
            value (float): ``pt`` value.

        Returns:
            UnitInch:
        """
        inst = super(UnitInch, cls).__new__(cls)
        return inst.__init__(UnitConvert.convert(num=value, frm=Length.PT, to=Length.IN))

    @classmethod
    def from_inch10(cls: Type[_TUnitInch], value: float) -> _TUnitInch:
        """
        Get instance from ``1/10th in`` (inch) value.

        Args:
            value (int): ``1/10th in`` value.

        Returns:
            UnitInch:
        """
        inst = super(UnitInch, cls).__new__(cls)
        return inst.__init__(round(UnitConvert.convert(num=value, frm=Length.IN10, to=Length.IN)))

    @classmethod
    def from_inch100(cls: Type[_TUnitInch], value: float) -> _TUnitInch:
        """
        Get instance from ``1/100th in`` (inch) value.

        Args:
            value (int): ``1/100th in`` value.

        Returns:
            UnitInch:
        """
        inst = super(UnitInch, cls).__new__(cls)
        return inst.__init__(round(UnitConvert.convert(num=value, frm=Length.IN100, to=Length.IN)))

    @classmethod
    def from_inch1000(cls: Type[_TUnitInch], value: float) -> _TUnitInch:
        """
        Get instance from ``1/1,000th in`` (inch) value.

        Args:
            value (int): ``1/1,000th in`` value.

        Returns:
            UnitInch:
        """
        inst = super(UnitInch, cls).__new__(cls)
        return inst.__init__(round(UnitConvert.convert(num=value, frm=Length.IN1000, to=Length.IN)))
