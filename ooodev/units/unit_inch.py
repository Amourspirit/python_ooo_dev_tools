from __future__ import annotations
from typing import TypeVar, Type
from dataclasses import dataclass
from ooodev.utils.data_type.base_float_value import BaseFloatValue
from .unit_convert import UnitConvert, UnitLength

_TUnitInch = TypeVar(name="_TUnitInch", bound="UnitInch")


@dataclass(unsafe_hash=True)
class UnitInch(BaseFloatValue):
    """
    Unit in ``inch`` units.

    Supports ``UnitObj`` protocol.

    See Also:
        :ref:`proto_unit_obj`
    """

    def __post_init__(self):
        if not isinstance(self.value, float):
            object.__setattr__(self, "value", float(self.value))

    def get_value_cm(self) -> float:
        """
        Gets instance value converted to ``cm`` units.

        Returns:
            int: Value in ``cm`` units.
        """
        return UnitConvert.convert(num=self.value, frm=UnitLength.IN, to=UnitLength.CM)

    def get_value_pt(self) -> float:
        """
        Gets instance value converted to Size in ``pt`` (points) units.

        Returns:
            int: Value in ``pt`` units.
        """
        return UnitConvert.convert(num=self.value, frm=UnitLength.IN, to=UnitLength.PT)

    def get_value_mm(self) -> float:
        """
        Gets instance value converted to Size in ``mm`` units.

        Returns:
            int: Value in ``mm`` units.
        """
        return UnitConvert.convert(num=self.value, frm=UnitLength.IN, to=UnitLength.MM)

    def get_value_mm100(self) -> int:
        """
        Gets instance value converted to Size in ``1/100th mm`` units.

        Returns:
            int: Value in ``1/100th mm`` units.
        """
        return round(UnitConvert.convert(num=self.value, frm=UnitLength.IN, to=UnitLength.MM100))

    def get_value_px(self) -> float:
        """
        Gets instance value in ``px`` (pixel) units.

        Returns:
            int: Value in ``px`` units.
        """
        return UnitConvert.convert(num=self.value, frm=UnitLength.IN, to=UnitLength.PX)

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
        inst.__init__(UnitConvert.convert(num=value, frm=UnitLength.MM, to=UnitLength.IN))
        return inst

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
        inst.__init__(UnitConvert.convert(num=value, frm=UnitLength.MM100, to=UnitLength.IN))
        return inst

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
        inst.__init__(UnitConvert.convert(num=value, frm=UnitLength.PT, to=UnitLength.IN))
        return inst

    @classmethod
    def from_px(cls: Type[_TUnitInch], value: float) -> _TUnitInch:
        """
        Get instance from ``px`` (pixel) value.

        Args:
            value (float): ``px`` value.

        Returns:
            UnitInch:
        """
        inst = super(UnitInch, cls).__new__(cls)
        inst.__init__(UnitConvert.convert(num=value, frm=UnitLength.PX, to=UnitLength.IN))
        return inst

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
        inst.__init__(UnitConvert.convert(num=value, frm=UnitLength.IN10, to=UnitLength.IN))
        return inst

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
        inst.__init__(UnitConvert.convert(num=value, frm=UnitLength.IN100, to=UnitLength.IN))
        return inst

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
        inst.__init__(UnitConvert.convert(num=value, frm=UnitLength.IN1000, to=UnitLength.IN))
        return inst

    @classmethod
    def from_cm(cls: Type[_TUnitInch], value: float) -> _TUnitInch:
        """
        Get instance from ``cm`` value.

        Args:
            value (float): ``cm`` value.

        Returns:
            UnitInch:
        """
        inst = super(UnitInch, cls).__new__(cls)
        inst.__init__(UnitConvert.convert(num=value, frm=UnitLength.CM, to=UnitLength.IN))
        return inst
