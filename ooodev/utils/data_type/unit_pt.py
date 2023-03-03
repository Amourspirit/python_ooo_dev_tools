from __future__ import annotations
from typing import TypeVar, Type
from dataclasses import dataclass
from ..validation import check
from .base_float_value import BaseFloatValue
from ..unit_convert import UnitConvert, Length

_TUnitPT = TypeVar(name="_TUnitPT", bound="UnitPT")


@dataclass(unsafe_hash=True)
class UnitPT(BaseFloatValue):
    """
    Represents a ``PT`` (points) value.

    Supports ``UnitObj`` protcol.

    See Also:
        :ref:`proto_unit_obj`
    """

    def __post_init__(self) -> None:
        if not isinstance(self.value, float):
            object.__setattr__(self, "value", float(self.value))
        check(
            self.value >= 0,
            f"{self}",
            f"Value of {self.value} is out of range. Value must be a positive number.",
        )

    def get_value_mm(self) -> float:
        """
        Gets instance value converted to ``mm`` units.

        Returns:
            int: Value in ``mm`` units.
        """
        return float(UnitConvert.convert(num=self.value, frm=Length.PT, to=Length.MM))

    def get_value_mm100(self) -> int:
        """
        Gets instance value converted to ``1/100th mm`` units.

        Returns:
            int: Value in ``1/100th mm`` units.
        """
        return round(UnitConvert.convert(num=self.value, frm=Length.PT, to=Length.MM100))

    def get_value_pt(self) -> float:
        """
        Gets instance value in ``pt`` (point) units.

        Returns:
            int: Value in ``pt`` units.
        """
        return self.value

    @classmethod
    def from_pt(cls: Type[_TUnitPT], value: float) -> _TUnitPT:
        """
        Get instance from ``pt`` (points) value.

        Args:
            value (float): ``pt`` value.

        Returns:
            UnitPT:
        """
        inst = super(UnitPT, cls).__new__(cls)
        return inst.__init__(value)

    @classmethod
    def from_px(cls: Type[_TUnitPT], value: float) -> _TUnitPT:
        """
        Get instance from ``px`` (pixel) value.

        Args:
            value (float): ``px`` value.

        Returns:
            UnitPT:
        """
        inst = super(UnitPT, cls).__new__(cls)
        return inst.__init__(UnitConvert.convert(num=value, frm=Length.PX, to=Length.PT))

    @classmethod
    def from_mm(cls: Type[_TUnitPT], value: float) -> _TUnitPT:
        """
        Get instance from ``mm`` value.

        Args:
            value (int): ``mm`` value.

        Returns:
            UnitPT:
        """
        inst = super(UnitPT, cls).__new__(cls)
        return inst.__init__(UnitConvert.convert(num=value, frm=Length.MM, to=Length.PT))

    @classmethod
    def from_mm10(cls: Type[_TUnitPT], value: float) -> _TUnitPT:
        """
        Get instance from ``1/10th mm`` value.

        Args:
            value (int): ``1/10th mm`` value.

        Returns:
            UnitPT:
        """
        inst = super(UnitPT, cls).__new__(cls)
        return inst.__init__(UnitConvert.convert(num=value, frm=Length.MM10, to=Length.PT))

    @classmethod
    def from_mm100(cls: Type[_TUnitPT], value: int) -> _TUnitPT:
        """
        Get instance from ``1/100th mm`` value.

        Args:
            value (int): ``1/100th mm`` value.

        Returns:
            UnitPT:
        """
        inst = super(UnitPT, cls).__new__(cls)
        return inst.__init__(UnitConvert.convert(num=value, frm=Length.MM100, to=Length.PT))

    @classmethod
    def from_inch(cls: Type[_TUnitPT], value: float) -> _TUnitPT:
        """
        Get instance from ``in`` (inch) value.

        Args:
            value (int): ``in`` value.

        Returns:
            UnitPT:
        """
        inst = super(UnitPT, cls).__new__(cls)
        return inst.__init__(UnitConvert.convert(num=value, frm=Length.IN, to=Length.PT))

    @classmethod
    def from_inch10(cls: Type[_TUnitPT], value: float) -> _TUnitPT:
        """
        Get instance from ``1/10th in`` (inch) value.

        Args:
            value (int): ``1/10th in`` value.

        Returns:
            UnitPT:
        """
        inst = super(UnitPT, cls).__new__(cls)
        return inst.__init__(UnitConvert.convert(num=value, frm=Length.IN10, to=Length.PT))

    @classmethod
    def from_inch100(cls: Type[_TUnitPT], value: float) -> _TUnitPT:
        """
        Get instance from ``1/100th in`` (inch) value.

        Args:
            value (int): ``1/100th in`` value.

        Returns:
            UnitPT:
        """
        inst = super(UnitPT, cls).__new__(cls)
        return inst.__init__(UnitConvert.convert(num=value, frm=Length.IN100, to=Length.PT))

    @classmethod
    def from_inch1000(cls: Type[_TUnitPT], value: int) -> _TUnitPT:
        """
        Get instance from ``1/1,000th in`` (inch) value.

        Args:
            value (int): ``1/1,000th in`` value.

        Returns:
            UnitPT:
        """
        inst = super(UnitPT, cls).__new__(cls)
        return inst.__init__(UnitConvert.convert(num=value, frm=Length.IN1000, to=Length.PT))
