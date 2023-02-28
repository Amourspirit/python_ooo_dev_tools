from __future__ import annotations
from typing import TypeVar, Type
from dataclasses import dataclass
from .base_float_value import BaseFloatValue
from ..unit_convert import UnitConvert, Length

_TUnitMM = TypeVar(name="_TUnitMM", bound="UnitMM")


@dataclass(unsafe_hash=True)
class UnitMM(BaseFloatValue):
    """
    Unit in ``mm`` units.

    Supports ``UnitObj`` protcol.

    See Also:
        :ref:`proto_unit_obj`
    """

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

    def get_value_mm(self) -> float:
        """
        Gets instance value converted to Size in ``mm`` units.

        Returns:
            int: Value in ``mm`` units.
        """
        return self.value

    def get_value_pt(self) -> float:
        """
        Gets instance value converted to Size in ``pt`` (points) units.

        Returns:
            int: Value in ``pt`` units.
        """
        return UnitConvert.convert(num=self.value, frm=Length.MM, to=Length.PT)

    def get_value_px(self) -> float:
        """
        Gets instance value in ``px`` (pixel) units.

        Returns:
            int: Value in ``px`` units.
        """
        return UnitConvert.convert(num=self.value, frm=Length.MM, to=Length.PX)

    @classmethod
    def from_mm(cls: Type[_TUnitMM], value: float) -> _TUnitMM:
        """
        Get instance from ``mm`` value.

        Args:
            value (int): ``mm`` value.

        Returns:
            UnitMM:
        """
        inst = super(UnitMM, cls).__new__(cls)
        return inst.__init__(value)

    @classmethod
    def from_mm10(cls: Type[_TUnitMM], value: float) -> _TUnitMM:
        """
        Get instance from ``1/10th mm`` value.

        Args:
            value (int): ``1/10th mm`` value.

        Returns:
            UnitMM:
        """
        inst = super(UnitMM, cls).__new__(cls)
        return inst.__init__(UnitConvert.convert(num=value, frm=Length.MM10, to=Length.MM))

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
        return inst.__init__(UnitConvert.convert_mm100_mm(value))

    @classmethod
    def from_pt(cls: Type[_TUnitMM], value: float) -> _TUnitMM:
        """
        Get instance from ``pt`` (points) value.

        Args:
            value (float): ``pt`` value.

        Returns:
            UnitMM:
        """
        inst = super(UnitMM, cls).__new__(cls)
        return inst.__init__(float(UnitConvert.convert(num=value, frm=Length.PT, to=Length.MM)))

    @classmethod
    def from_px(cls: Type[_TUnitMM], value: float) -> _TUnitMM:
        """
        Get instance from ``px`` (pixel) value.

        Args:
            value (float): ``px`` value.

        Returns:
            UnitMM:
        """
        inst = super(UnitMM, cls).__new__(cls)
        return inst.__init__(UnitConvert.convert(num=value, frm=Length.PX, to=Length.MM))

    @classmethod
    def from_inch(cls: Type[_TUnitMM], value: float) -> _TUnitMM:
        """
        Get instance from ``in`` (inch) value.

        Args:
            value (int): ``in`` value.

        Returns:
            UnitMM:
        """
        inst = super(UnitMM, cls).__new__(cls)
        return inst.__init__(UnitConvert.convert(num=value, frm=Length.IN, to=Length.MM))

    @classmethod
    def from_inch10(cls: Type[_TUnitMM], value: float) -> _TUnitMM:
        """
        Get instance from ``1/10th in`` (inch) value.

        Args:
            value (int): ```/10th in`` value.

        Returns:
            UnitMM:
        """
        inst = super(UnitMM, cls).__new__(cls)
        return inst.__init__(UnitConvert.convert(num=value, frm=Length.IN10, to=Length.MM))

    @classmethod
    def from_inch100(cls: Type[_TUnitMM], value: float) -> _TUnitMM:
        """
        Get instance from ``1/100th in`` (inch) value.

        Args:
            value (int): ``1/100th in`` value.

        Returns:
            UnitMM:
        """
        inst = super(UnitMM, cls).__new__(cls)
        return inst.__init__(UnitConvert.convert(num=value, frm=Length.IN100, to=Length.MM))

    @classmethod
    def from_inch1000(cls: Type[_TUnitMM], value: int) -> _TUnitMM:
        """
        Get instance from ``1/1,000th in`` (inch) value.

        Args:
            value (int): ``1/1,000th in`` value.

        Returns:
            UnitMM:
        """
        inst = super(UnitMM, cls).__new__(cls)
        return inst.__init__(UnitConvert.convert(num=value, frm=Length.IN1000, to=Length.MM))
