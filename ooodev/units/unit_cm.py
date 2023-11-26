from __future__ import annotations
from typing import TypeVar, Type
from dataclasses import dataclass
from ooodev.utils.data_type.base_float_value import BaseFloatValue
from .unit_convert import UnitConvert, UnitLength

_TUnitCM = TypeVar(name="_TUnitCM", bound="UnitCM")


@dataclass(unsafe_hash=True)
class UnitCM(BaseFloatValue):
    """
    Unit in ``cm`` units.

    Supports ``UnitT`` protocol.

    See Also:
        :ref:`proto_unit_obj`

    .. versionadded:: 0.9.4
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
        return self.value

    def get_value_mm(self) -> float:
        """
        Gets instance value converted to ``mm`` units.

        Returns:
            int: Value in ``mm`` units.
        """
        return UnitConvert.convert(num=self.value, frm=UnitLength.CM, to=UnitLength.MM)

    def get_value_mm100(self) -> int:
        """
        Gets instance value converted to ``1/100th mm`` units.

        Returns:
            int: Value in ``1/100th mm`` units.
        """
        return round(UnitConvert.convert(num=self.value, frm=UnitLength.CM, to=UnitLength.MM100))

    def get_value_pt(self) -> float:
        """
        Gets instance value converted to ``pt`` (points) units.

        Returns:
            int: Value in ``pt`` units.
        """
        return UnitConvert.convert(num=self.value, frm=UnitLength.CM, to=UnitLength.PT)

    def get_value_px(self) -> float:
        """
        Gets instance value in ``px`` (pixel) units.

        Returns:
            int: Value in ``px`` units.
        """
        return UnitConvert.convert(num=self.value, frm=UnitLength.CM, to=UnitLength.PX)

    @classmethod
    def from_pt(cls: Type[_TUnitCM], value: float) -> _TUnitCM:
        """
        Get instance from ``pt`` (points) value.

        Args:
            value (float): ``pt`` value.

        Returns:
            UnitCM:
        """
        inst = super(UnitCM, cls).__new__(cls)
        inst.__init__(float(UnitConvert.convert(num=value, frm=UnitLength.PT, to=UnitLength.CM)))
        return inst

    @classmethod
    def from_px(cls: Type[_TUnitCM], value: float) -> _TUnitCM:
        """
        Get instance from ``px`` (pixel) value.

        Args:
            value (float): ``px`` value.

        Returns:
            UnitCM:
        """
        inst = super(UnitCM, cls).__new__(cls)
        inst.__init__(UnitConvert.convert(num=value, frm=UnitLength.PX, to=UnitLength.CM))
        return inst

    @classmethod
    def from_mm(cls: Type[_TUnitCM], value: float) -> _TUnitCM:
        """
        Get instance from ``mm`` value.

        Args:
            value (int): ``mm`` value.

        Returns:
            UnitCM:
        """
        inst = super(UnitCM, cls).__new__(cls)
        inst.__init__(UnitConvert.convert(num=value, frm=UnitLength.MM, to=UnitLength.CM))
        return inst

    @classmethod
    def from_mm10(cls: Type[_TUnitCM], value: int) -> _TUnitCM:
        """
        Get instance from ``1/10th mm`` value.

        Args:
            value (int): ``1/10th mm`` value.

        Returns:
            UnitCM:
        """
        inst = super(UnitCM, cls).__new__(cls)
        inst.__init__(UnitConvert.convert(num=value, frm=UnitLength.MM10, to=UnitLength.CM))
        return inst

    @classmethod
    def from_mm100(cls: Type[_TUnitCM], value: int) -> _TUnitCM:
        """
        Get instance from ``1/100th mm`` value.

        Args:
            value (int): ``1/100th mm`` value.

        Returns:
            UnitCM:
        """
        inst = super(UnitCM, cls).__new__(cls)
        inst.__init__(UnitConvert.convert(num=value, frm=UnitLength.MM100, to=UnitLength.CM))
        return inst

    @classmethod
    def from_inch(cls: Type[_TUnitCM], value: float) -> _TUnitCM:
        """
        Get instance from ``in`` (inch) value.

        Args:
            value (int): ``in`` value.

        Returns:
            UnitCM:
        """
        inst = super(UnitCM, cls).__new__(cls)
        inst.__init__(UnitConvert.convert(num=value, frm=UnitLength.IN, to=UnitLength.CM))
        return inst

    @classmethod
    def from_inch10(cls: Type[_TUnitCM], value: float) -> _TUnitCM:
        """
        Get instance from ``1/10th in`` (inch) value.

        Args:
            value (int): ```/10th in`` value.

        Returns:
            UnitCM:
        """
        inst = super(UnitCM, cls).__new__(cls)
        inst.__init__(UnitConvert.convert(num=value, frm=UnitLength.IN10, to=UnitLength.CM))
        return inst

    @classmethod
    def from_inch100(cls: Type[_TUnitCM], value: float) -> _TUnitCM:
        """
        Get instance from ``1/100th in`` (inch) value.

        Args:
            value (int): ``1/100th in`` value.

        Returns:
            UnitCM:
        """
        inst = super(UnitCM, cls).__new__(cls)
        inst.__init__(UnitConvert.convert(num=value, frm=UnitLength.IN100, to=UnitLength.CM))
        return inst

    @classmethod
    def from_inch1000(cls: Type[_TUnitCM], value: int) -> _TUnitCM:
        """
        Get instance from ``1/1,000th in`` (inch) value.

        Args:
            value (int): ``1/1,000th in`` value.

        Returns:
            UnitCM:
        """
        inst = super(UnitCM, cls).__new__(cls)
        inst.__init__(UnitConvert.convert(num=value, frm=UnitLength.IN1000, to=UnitLength.CM))
        return inst

    @classmethod
    def from_cm(cls: Type[_TUnitCM], value: float) -> _TUnitCM:
        """
        Get instance from ``cm`` value.

        Args:
            value (float): ``cm`` value.

        Returns:
            UnitCM:
        """
        inst = super(UnitCM, cls).__new__(cls)
        inst.__init__(value)
        return inst
