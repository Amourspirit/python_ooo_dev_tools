from __future__ import annotations
from typing import TypeVar, Type
from dataclasses import dataclass
from ..validation import check
from ..decorator import enforce
from .base_int_value import BaseIntValue
from ..unit_convert import UnitConvert, Length

_TUnitInch1000 = TypeVar(name="_TUnitInch1000", bound="UnitInch1000")

# Note that from __future__ import annotations converts annotations to string.
# this means that @enforce.enforce_types will see string as type. This is fine in
# most cases. Especially for built in types.
@enforce.enforce_types
@dataclass(unsafe_hash=True)
class UnitInch1000(BaseIntValue):
    """
    Represents ``1/1,000th in`` units.

    Supports ``UnitObj`` protcol.

    See Also:
        :ref:`proto_unit_obj`
    """

    def _from_int(self, value: int) -> _TUnitInch1000:
        inst = super(UnitInch1000, self.__class__).__new__(self.__class__)
        return inst.__init__(value)

    def __eq__(self, other: object) -> bool:
        # for some reason BaseIntValue __eq__ is not picked up.
        # I suspect this is due to this class being a dataclass.
        try:
            i = int(other)
            return i == self.value
        except Exception as e:
            return False

    def get_value_mm(self) -> float:
        """
        Gets instance value converted to ``mm`` units.

        Returns:
            int: Value in ``mm`` units.
        """
        return UnitConvert.convert(num=self.value, frm=Length.IN1000, to=Length.MM)

    def get_value_mm100(self) -> int:
        """
        Gets instance value in ``1/100th mm`` units.

        Returns:
            int: Value in ``1/100th mm`` units.
        """
        return round(UnitConvert.convert(num=self.value, frm=Length.IN1000, to=Length.MM100))

    def get_value_pt(self) -> float:
        """
        Gets instance value converted to ``pt`` (point) units.

        Returns:
            int: Value in ``pt`` units.
        """
        return UnitConvert.convert(num=self.value, frm=Length.IN1000, to=Length.PT)

    def get_value_px(self) -> float:
        """
        Gets instance value in ``px`` (pixel) units.

        Returns:
            int: Value in ``px`` units.
        """
        return UnitConvert.convert(num=self.value, frm=Length.IN1000, to=Length.PX)

    @classmethod
    def from_mm(cls: Type[_TUnitInch1000], value: float) -> _TUnitInch1000:
        """
        Get instance from ``mm`` value.

        Args:
            value (int): ``mm`` value.

        Returns:
            UnitInch1000:
        """
        inst = super(UnitInch1000, cls).__new__(cls)
        inst.__init__(round(UnitConvert.convert(num=value, frm=Length.MM, to=Length.IN1000)))
        return inst

    @classmethod
    def from_pt(cls: Type[_TUnitInch1000], value: float) -> _TUnitInch1000:
        """
        Get instance from ``pt`` (points) value.

        Args:
            value (int): ``pt`` value.

        Returns:
            UnitInch1000:
        """
        inst = super(UnitInch1000, cls).__new__(cls)
        inst.__init__(round(UnitConvert.convert(num=value, frm=Length.PT, to=Length.IN1000)))
        return inst

    @classmethod
    def from_px(cls: Type[_TUnitInch1000], value: float) -> _TUnitInch1000:
        """
        Get instance from ``px`` (pixel) value.

        Args:
            value (float): ``px`` value.

        Returns:
            UnitInch1000:
        """
        inst = super(UnitInch1000, cls).__new__(cls)
        inst.__init__(round(UnitConvert.convert(num=value, frm=Length.PX, to=Length.IN1000)))
        return inst

    @classmethod
    def from_inch(cls: Type[_TUnitInch1000], value: float) -> _TUnitInch1000:
        """
        Get instance from ``in`` (inch) value.

        Args:
            value (int): ``in`` value.

        Returns:
            UnitInch1000:
        """
        inst = super(UnitInch1000, cls).__new__(cls)
        inst.__init__(round(UnitConvert.convert(num=value, frm=Length.IN, to=Length.IN1000)))
        return inst

    @classmethod
    def from_inch10(cls: Type[_TUnitInch1000], value: float) -> _TUnitInch1000:
        """
        Get instance from ``1/10th in`` (inch) value.

        Args:
            value (int): ```/10th in`` value.

        Returns:
            UnitInch1000:
        """
        inst = super(UnitInch1000, cls).__new__(cls)
        inst.__init__(round(UnitConvert.convert(num=value, frm=Length.IN10, to=Length.IN1000)))
        return inst

    @classmethod
    def from_inch100(cls: Type[_TUnitInch1000], value: float) -> _TUnitInch1000:
        """
        Get instance from ``1/100th in`` (inch) value.

        Args:
            value (int): ``1/100th in`` value.

        Returns:
            UnitInch1000:
        """
        inst = super(UnitInch1000, cls).__new__(cls)
        inst.__init__(round(UnitConvert.convert(num=value, frm=Length.IN100, to=Length.IN1000)))
        return inst

    @classmethod
    def from_inch1000(cls: Type[_TUnitInch1000], value: int) -> _TUnitInch1000:
        """
        Get instance from ``1/1,000th in`` (inch) value.

        Args:
            value (int): ``1/1,000th in`` value.

        Returns:
            UnitInch1000:
        """
        inst = super(UnitInch1000, cls).__new__(cls)
        inst.__init__(value)
        return inst
