from __future__ import annotations
import contextlib
from typing import TypeVar, Type, TYPE_CHECKING
from dataclasses import dataclass
from ooodev.utils.data_type.base_float_value import BaseFloatValue
from .unit_convert import UnitConvert, UnitLength

if TYPE_CHECKING:
    from ooodev.units import UnitT

_TUnitInch100 = TypeVar(name="_TUnitInch100", bound="UnitInch100")


@dataclass(unsafe_hash=True)
class UnitInch100(BaseFloatValue):
    """
    Unit in ``1/100th in`` units.

    Supports ``UnitT`` protocol.

    See Also:
        :ref:`proto_unit_obj`
    """

    def __post_init__(self):
        if not isinstance(self.value, float):
            object.__setattr__(self, "value", float(self.value))

    # region math and comparison
    def __int__(self) -> int:
        return round(self.value)

    def __eq__(self, other: object) -> bool:
        if isinstance(other, UnitInch100):
            return self.almost_equal(other.value)
        if hasattr(other, "get_value_inch100"):
            oth_val = other.get_value_inch100()  # type: ignore
            return self.almost_equal(oth_val)
        if hasattr(other, "get_value_mm100"):
            return self.get_value_mm100() == other.get_value_mm100()  # type: ignore
        with contextlib.suppress(Exception):
            return self.almost_equal(float(other))  # type: ignore
        return False

    def __add__(self, other: object) -> UnitInch100:
        if isinstance(other, UnitInch100):
            return self.from_inch100(self.value + other.value)
        if hasattr(other, "get_value_inch100"):
            oth_val = other.get_value_inch100()  # type: ignore
            return self.from_inch100(self.value + oth_val)
        if hasattr(other, "get_value_mm100"):
            oth_val = other.get_value_mm100()  # type: ignore
            oth_val_inch100 = UnitConvert.convert(num=oth_val, frm=UnitLength.MM100, to=UnitLength.IN100)
            return self.from_inch100(self.value + oth_val_inch100)  # type: ignore

        if isinstance(other, (int, float)):
            return self.from_inch100(self.value + other)  # type: ignore
        return NotImplemented

    def __radd__(self, other: object) -> UnitInch100:
        return self if other == 0 else self.__add__(other)

    def __sub__(self, other: object) -> UnitInch100:
        if isinstance(other, UnitInch100):
            return self.from_inch100(self.value - other.value)
        if hasattr(other, "get_value_inch100"):
            oth_val = other.get_value_inch100()  # type: ignore
            return self.from_inch100(self.value - oth_val)
        if hasattr(other, "get_value_mm100"):
            oth_val = other.get_value_mm100()  # type: ignore
            oth_val_inch100 = UnitConvert.convert(num=oth_val, frm=UnitLength.MM100, to=UnitLength.IN100)
            return self.from_inch100(self.value - oth_val_inch100)  # type: ignore

        if isinstance(other, (int, float)):
            return self.from_inch100(self.value - other)  # type: ignore
        return NotImplemented

    def __rsub__(self, other: object) -> UnitInch100:
        if isinstance(other, (int, float)):
            return self.from_inch100(other - self.value)  # type: ignore
        return NotImplemented

    def __mul__(self, other: object) -> UnitInch100:
        if isinstance(other, UnitInch100):
            return self.from_inch100(self.value * other.value)
        if hasattr(other, "get_value_inch100"):
            oth_val = other.get_value_inch100()  # type: ignore
            return self.from_inch100(self.value * oth_val)
        if hasattr(other, "get_value_mm100"):
            oth_val = other.get_value_mm100()  # type: ignore
            oth_val_inch100 = UnitConvert.convert(num=oth_val, frm=UnitLength.MM100, to=UnitLength.IN100)
            return self.from_inch100(self.value * oth_val_inch100)  # type: ignore

        if isinstance(other, (int, float)):
            return self.from_inch100(self.value * other)  # type: ignore

        return NotImplemented

    def __rmul__(self, other: int) -> UnitInch100:
        return self if other == 0 else self.__mul__(other)

    def __truediv__(self, other: object) -> UnitInch100:
        if isinstance(other, UnitInch100):
            if other.value == 0:
                raise ZeroDivisionError
            return self.from_inch100(self.value / other.value)
        if hasattr(other, "get_value_inch100"):
            oth_val = other.get_value_inch100()  # type: ignore
            if oth_val == 0:
                raise ZeroDivisionError
            return self.from_inch100(self.value / oth_val)
        if hasattr(other, "get_value_mm100"):
            oth_val = other.get_value_mm100()  # type: ignore
            oth_val_inch100 = UnitConvert.convert(num=oth_val, frm=UnitLength.MM100, to=UnitLength.IN100)
            if oth_val_inch100 == 0:
                raise ZeroDivisionError
            return self.from_inch100(self.value / oth_val_inch100)  # type: ignore
        if isinstance(other, (int, float)):
            if other == 0:
                raise ZeroDivisionError
            return self.from_inch100(self.value / other)  # type: ignore
        return NotImplemented

    def __rtruediv__(self, other: object) -> UnitInch100:
        if isinstance(other, (int, float)):
            if self.value == 0:
                raise ZeroDivisionError
            return self.from_inch100(other / self.value)  # type: ignore
        return NotImplemented

    def __abs__(self) -> float:
        return abs(self.value)

    def __lt__(self, other: object) -> bool:
        if isinstance(other, UnitInch100):
            return self.value < other.value
        if hasattr(other, "get_value_inch100"):
            oth_val = other.get_value_inch100()  # type: ignore
            return self.value < oth_val
        if hasattr(other, "get_value_mm100"):
            return self.get_value_mm100() < other.get_value_mm100()  # type: ignore
        with contextlib.suppress(Exception):
            return self.value < float(other)  # type: ignore
        return False

    def __le__(self, other: object) -> bool:
        if isinstance(other, UnitInch100):
            if self.almost_equal(other.value):
                return True
            return self.value < other.value
        if hasattr(other, "get_value_inch100"):
            oth_val = other.get_value_inch100()  # type: ignore
            if self.almost_equal(oth_val):
                return True
            return self.value < oth_val
        if hasattr(other, "get_value_mm100"):
            return self.get_value_mm100() <= other.get_value_mm100()  # type: ignore
        with contextlib.suppress(Exception):
            oth_val = float(other)  # type: ignore
            if self.almost_equal(oth_val):
                return True
            return self.value < oth_val
        return False

    def __gt__(self, other: object) -> bool:
        if isinstance(other, UnitInch100):
            return self.value > other.value
        if hasattr(other, "get_value_inch100"):
            oth_val = other.get_value_inch100()  # type: ignore
            return self.value > oth_val
        if hasattr(other, "get_value_mm100"):
            return self.get_value_mm100() > other.get_value_mm100()  # type: ignore
        with contextlib.suppress(Exception):
            return self.value > float(other)  # type: ignore
        return False

    def __ge__(self, other: object) -> bool:
        if isinstance(other, UnitInch100):
            if self.almost_equal(other.value):
                return True
            return self.value > other.value
        if hasattr(other, "get_value_inch100"):
            oth_val = other.get_value_inch100()  # type: ignore
            if self.almost_equal(oth_val):
                return True
            return self.value > oth_val
        if hasattr(other, "get_value_mm100"):
            return self.get_value_mm100() >= other.get_value_mm100()  # type: ignore
        with contextlib.suppress(Exception):
            oth_val = float(other)  # type: ignore
            if self.almost_equal(oth_val):
                return True
            return self.value > oth_val
        return False

    # endregion math and comparison

    def get_value_cm(self) -> float:
        """
        Gets instance value converted to ``cm`` units.

        Returns:
            int: Value in ``cm`` units.
        """
        return UnitConvert.convert(num=self.value, frm=UnitLength.IN100, to=UnitLength.CM)

    def get_value_mm(self) -> float:
        """
        Gets instance value converted to Size in ``mm`` units.

        Returns:
            int: Value in ``mm`` units.
        """
        return UnitConvert.convert(num=self.value, frm=UnitLength.IN100, to=UnitLength.MM)

    def get_value_mm100(self) -> int:
        """
        Gets instance value converted to Size in ``1/100th mm`` units.

        Returns:
            int: Value in ``1/100th mm`` units.
        """
        return round(UnitConvert.convert(num=self.value, frm=UnitLength.IN100, to=UnitLength.MM100))

    def get_value_inch1000(self) -> int:
        """
        Gets instance value in ``1/1000th inch`` units.

        Returns:
            int: Value in ``1/1000th inch`` units.
        """
        return round(UnitConvert.convert(num=self.value, frm=UnitLength.IN100, to=UnitLength.IN1000))

    def get_value_px(self) -> float:
        """
        Gets instance value in ``px`` (pixel) units.

        Returns:
            int: Value in ``px`` units.
        """
        return UnitConvert.convert(num=self.value, frm=UnitLength.IN100, to=UnitLength.PX)

    def get_value_pt(self) -> float:
        """
        Gets instance value converted to Size in ``pt`` (points) units.

        Returns:
            int: Value in ``pt`` units.
        """
        return UnitConvert.convert(num=self.value, frm=UnitLength.IN100, to=UnitLength.PT)

    @classmethod
    def from_mm100(cls: Type[_TUnitInch100], value: int) -> _TUnitInch100:
        """
        Get instance from ``1/100th mm`` value.

        Args:
            value (int): ``1/100th mm`` value.

        Returns:
            UnitInch100:
        """
        inst = super(UnitInch100, cls).__new__(cls)  # type: ignore
        inst.__init__(UnitConvert.convert(num=value, frm=UnitLength.MM100, to=UnitLength.IN100))
        return inst

    @classmethod
    def from_pt(cls: Type[_TUnitInch100], value: float) -> _TUnitInch100:
        """
        Get instance from ``pt`` (points) value.

        Args:
            value (float): ``pt`` value.

        Returns:
            UnitInch100:
        """
        inst = super(UnitInch100, cls).__new__(cls)  # type: ignore
        inst.__init__(float(UnitConvert.convert(num=value, frm=UnitLength.PT, to=UnitLength.IN100)))
        return inst

    @classmethod
    def from_px(cls: Type[_TUnitInch100], value: float) -> _TUnitInch100:
        """
        Get instance from ``px`` (pixel) value.

        Args:
            value (float): ``px`` value.

        Returns:
            UnitInch100:
        """
        inst = super(UnitInch100, cls).__new__(cls)  # type: ignore
        inst.__init__(UnitConvert.convert(num=value, frm=UnitLength.PX, to=UnitLength.IN100))
        return inst

    @classmethod
    def from_inch(cls: Type[_TUnitInch100], value: float) -> _TUnitInch100:
        """
        Get instance from ``in`` (inch) value.

        Args:
            value (float): ``in`` value.

        Returns:
            UnitInch100:
        """
        inst = super(UnitInch100, cls).__new__(cls)  # type: ignore
        inst.__init__(UnitConvert.convert(num=value, frm=UnitLength.IN, to=UnitLength.IN100))
        return inst

    from_in = from_inch

    @classmethod
    def from_inch10(cls: Type[_TUnitInch100], value: float) -> _TUnitInch100:
        """
        Get instance from ``1/10th in`` (inch) value.

        Args:
            value (int): ``1/10th in`` value.

        Returns:
            UnitInch100:
        """
        inst = super(UnitInch100, cls).__new__(cls)  # type: ignore
        inst.__init__(UnitConvert.convert(num=value, frm=UnitLength.IN10, to=UnitLength.IN100))
        return inst

    @classmethod
    def from_inch100(cls: Type[_TUnitInch100], value: float) -> _TUnitInch100:
        """
        Get instance from ``1/10th in`` (inch) value.

        Args:
            value (int): ``1/10th in`` value.

        Returns:
            UnitInch100:
        """
        inst = super(UnitInch100, cls).__new__(cls)  # type: ignore
        inst.__init__(value)
        return inst

    @classmethod
    def from_inch1000(cls: Type[_TUnitInch100], value: int) -> _TUnitInch100:
        """
        Get instance from ``1/1,000th in`` (inch) value.

        Args:
            value (int): ``1/1,000th in`` value.

        Returns:
            UnitInch100:
        """
        inst = super(UnitInch100, cls).__new__(cls)  # type: ignore
        inst.__init__(UnitConvert.convert(num=value, frm=UnitLength.IN1000, to=UnitLength.IN100))
        return inst

    @classmethod
    def from_cm(cls: Type[_TUnitInch100], value: float) -> _TUnitInch100:
        """
        Get instance from ``cm`` value.

        Args:
            value (float): ``cm`` value.

        Returns:
            UnitInch100:
        """
        inst = super(UnitInch100, cls).__new__(cls)  # type: ignore
        inst.__init__(UnitConvert.convert(num=value, frm=UnitLength.CM, to=UnitLength.IN100))
        return inst

    @classmethod
    def from_unit_val(cls: Type[_TUnitInch100], value: UnitT | float) -> _TUnitInch100:
        """
        Get instance from ``UnitT`` or float value.

        Args:
            value (UnitT, float): ``UnitT`` or float value. If float then it is assumed to be in ``inch100`` units.

        Returns:
            UnitInch100:
        """
        try:
            if hasattr(value, "get_value_inch100"):
                return cls.from_inch100(value.get_value_inch100())  # type: ignore

            unit_100 = value.get_value_mm100()  # type: ignore
            return cls.from_mm100(unit_100)
        except AttributeError:
            return cls.from_inch100(float(value))  # type: ignore
