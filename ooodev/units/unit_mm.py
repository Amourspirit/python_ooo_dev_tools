from __future__ import annotations
import contextlib
from typing import TypeVar, Type, TYPE_CHECKING
from dataclasses import dataclass
from ooodev.utils.data_type.base_float_value import BaseFloatValue
from .unit_convert import UnitConvert, UnitLength

if TYPE_CHECKING:
    from ooodev.units import UnitT

_TUnitMM = TypeVar(name="_TUnitMM", bound="UnitMM")


@dataclass(unsafe_hash=True)
class UnitMM(BaseFloatValue):
    """
    Unit in ``mm`` units.

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
        if isinstance(other, UnitMM):
            return self.almost_equal(other.value)
        if hasattr(other, "get_value_mm"):
            oth_val = other.get_value_mm()  # type: ignore
            return self.almost_equal(oth_val)
        if hasattr(other, "get_value_mm100"):
            return self.get_value_mm100() == other.get_value_mm100()  # type: ignore
        with contextlib.suppress(Exception):
            return self.almost_equal(float(other))  # type: ignore
        return False

    def __add__(self, other: object) -> UnitMM:
        if isinstance(other, UnitMM):
            return self.from_mm(self.value + other.value)
        if hasattr(other, "get_value_mm"):
            oth_val = other.get_value_mm()  # type: ignore
            return self.from_mm(self.value + oth_val)
        if hasattr(other, "get_value_mm100"):
            oth_val = other.get_value_mm100()  # type: ignore
            oth_val_mm = UnitConvert.convert(num=oth_val, frm=UnitLength.MM100, to=UnitLength.MM)
            return self.from_mm(self.value + oth_val_mm)  # type: ignore

        if isinstance(other, (int, float)):
            return self.from_mm(self.value + other)  # type: ignore
        return NotImplemented

    def __radd__(self, other: object) -> UnitMM:
        return self if other == 0 else self.__add__(other)

    def __sub__(self, other: object) -> UnitMM:
        if isinstance(other, UnitMM):
            return self.from_mm(self.value - other.value)
        if hasattr(other, "get_value_mm"):
            oth_val = other.get_value_mm()  # type: ignore
            return self.from_mm(self.value - oth_val)
        if hasattr(other, "get_value_mm100"):
            oth_val = other.get_value_mm100()  # type: ignore
            oth_val_mm = UnitConvert.convert(num=oth_val, frm=UnitLength.MM100, to=UnitLength.MM)
            return self.from_mm(self.value - oth_val_mm)  # type: ignore

        if isinstance(other, (int, float)):
            return self.from_mm(self.value - other)  # type: ignore
        return NotImplemented

    def __rsub__(self, other: object) -> UnitMM:
        if isinstance(other, (int, float)):
            return self.from_mm(other - self.value)  # type: ignore
        return NotImplemented

    def __mul__(self, other: object) -> UnitMM:
        if isinstance(other, UnitMM):
            return self.from_mm(self.value * other.value)
        if hasattr(other, "get_value_mm"):
            oth_val = other.get_value_mm()  # type: ignore
            return self.from_mm(self.value * oth_val)
        if hasattr(other, "get_value_mm100"):
            oth_val = other.get_value_mm100()  # type: ignore
            oth_val_mm = UnitConvert.convert(num=oth_val, frm=UnitLength.MM100, to=UnitLength.MM)
            return self.from_mm(self.value * oth_val_mm)

        if isinstance(other, (int, float)):
            return self.from_mm(self.value * other)  # type: ignore

        return NotImplemented

    def __rmul__(self, other: int) -> UnitMM:
        return self if other == 0 else self.__mul__(other)

    def __truediv__(self, other: object) -> UnitMM:
        if isinstance(other, UnitMM):
            if other.value == 0:
                raise ZeroDivisionError
            return self.from_mm(self.value / other.value)
        if hasattr(other, "get_value_mm"):
            oth_val = other.get_value_mm()  # type: ignore
            if oth_val == 0:
                raise ZeroDivisionError
            return self.from_mm(self.value / oth_val)
        if hasattr(other, "get_value_mm100"):
            oth_val = other.get_value_mm100()  # type: ignore
            oth_val_mm = UnitConvert.convert(num=oth_val, frm=UnitLength.MM100, to=UnitLength.MM)
            if oth_val_mm == 0:
                raise ZeroDivisionError
            return self.from_mm(self.value / oth_val_mm)  # type: ignore
        if isinstance(other, (int, float)):
            if other == 0:
                raise ZeroDivisionError
            return self.from_mm(self.value / other)  # type: ignore
        return NotImplemented

    def __rtruediv__(self, other: object) -> UnitMM:
        if isinstance(other, (int, float)):
            if self.value == 0:
                raise ZeroDivisionError
            return self.from_mm(other / self.value)  # type: ignore
        return NotImplemented

    def __abs__(self) -> float:
        return abs(self.value)

    def __lt__(self, other: object) -> bool:
        if isinstance(other, UnitMM):
            return self.value < other.value
        if hasattr(other, "get_value_mm"):
            oth_val = other.get_value_mm()  # type: ignore
            return self.value < oth_val
        if hasattr(other, "get_value_mm100"):
            return self.get_value_mm100() < other.get_value_mm100()  # type: ignore
        with contextlib.suppress(Exception):
            return self.value < float(other)  # type: ignore
        return False

    def __le__(self, other: object) -> bool:
        if isinstance(other, UnitMM):
            if self.almost_equal(other.value):
                return True
            return self.value < other.value
        if hasattr(other, "get_value_mm"):
            oth_val = other.get_value_mm()  # type: ignore
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
        if isinstance(other, UnitMM):
            return self.value > other.value
        if hasattr(other, "get_value_mm"):
            oth_val = other.get_value_mm()  # type: ignore
            return self.value > oth_val
        if hasattr(other, "get_value_mm100"):
            return self.get_value_mm100() > other.get_value_mm100()  # type: ignore
        with contextlib.suppress(Exception):
            return self.value > float(other)  # type: ignore
        return False

    def __ge__(self, other: object) -> bool:
        if isinstance(other, UnitMM):
            if self.almost_equal(other.value):
                return True
            return self.value > other.value
        if hasattr(other, "get_value_mm"):
            oth_val = other.get_value_mm()  # type: ignore
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
            float: Value in ``cm`` units.
        """
        return UnitConvert.convert(num=self.value, frm=UnitLength.MM, to=UnitLength.CM)

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
            float: Value in ``mm`` units.
        """
        return self.value

    def get_value_pt(self) -> float:
        """
        Gets instance value converted to Size in ``pt`` (points) units.

        Returns:
            float: Value in ``pt`` units.
        """
        return UnitConvert.convert(num=self.value, frm=UnitLength.MM, to=UnitLength.PT)

    def get_value_px(self) -> float:
        """
        Gets instance value in ``px`` (pixel) units.

        Returns:
            float: Value in ``px`` units.
        """
        return UnitConvert.convert(num=self.value, frm=UnitLength.MM, to=UnitLength.PX)

    def get_value_inch(self) -> float:
        """
        Gets instance value in ``in`` (inch) units.

        Returns:
            float: Value in ``in`` units.
        """
        return UnitConvert.convert(num=self.value, frm=UnitLength.MM, to=UnitLength.IN)

    def get_value_inch100(self) -> int:
        """
        Gets instance value in ``1/100th inch`` units.

        Returns:
            int: Value in ``1/100th inch`` units.
        """
        return round(UnitConvert.convert(num=self.value, frm=UnitLength.MM, to=UnitLength.IN100))

    def get_value_inch10(self) -> int:
        """
        Gets instance value in ``1/10th inch`` units.

        Returns:
            int: Value in ``1/10th inch`` units.
        """
        return round(UnitConvert.convert(num=self.value, frm=UnitLength.MM, to=UnitLength.IN10))

    def get_value_inch1000(self) -> int:
        """
        Gets instance value in ``1/1000th inch`` units.

        Returns:
            int: Value in ``1/100th inch`` units.
        """
        return round(UnitConvert.convert(num=self.value, frm=UnitLength.MM, to=UnitLength.IN1000))

    @classmethod
    def from_mm(cls: Type[_TUnitMM], value: float) -> _TUnitMM:
        """
        Get instance from ``mm`` value.

        Args:
            value (int): ``mm`` value.

        Returns:
            UnitMM:
        """
        inst = super(UnitMM, cls).__new__(cls)  # type: ignore
        inst.__init__(value)
        return inst

    @classmethod
    def from_mm10(cls: Type[_TUnitMM], value: float) -> _TUnitMM:
        """
        Get instance from ``1/10th mm`` value.

        Args:
            value (int): ``1/10th mm`` value.

        Returns:
            UnitMM:
        """
        inst = super(UnitMM, cls).__new__(cls)  # type: ignore
        inst.__init__(UnitConvert.convert(num=value, frm=UnitLength.MM10, to=UnitLength.MM))
        return inst

    @classmethod
    def from_mm100(cls: Type[_TUnitMM], value: int) -> _TUnitMM:
        """
        Get instance from ``1/100th mm`` value.

        Args:
            value (int): ``1/100th mm`` value.

        Returns:
            UnitMM:
        """
        inst = super(UnitMM, cls).__new__(cls)  # type: ignore
        inst.__init__(UnitConvert.convert_mm100_mm(value))
        return inst

    @classmethod
    def from_pt(cls: Type[_TUnitMM], value: float) -> _TUnitMM:
        """
        Get instance from ``pt`` (points) value.

        Args:
            value (float): ``pt`` value.

        Returns:
            UnitMM:
        """
        inst = super(UnitMM, cls).__new__(cls)  # type: ignore
        inst.__init__(float(UnitConvert.convert(num=value, frm=UnitLength.PT, to=UnitLength.MM)))
        return inst

    @classmethod
    def from_px(cls: Type[_TUnitMM], value: float) -> _TUnitMM:
        """
        Get instance from ``px`` (pixel) value.

        Args:
            value (float): ``px`` value.

        Returns:
            UnitMM:
        """
        inst = super(UnitMM, cls).__new__(cls)  # type: ignore
        inst.__init__(UnitConvert.convert(num=value, frm=UnitLength.PX, to=UnitLength.MM))
        return inst

    @classmethod
    def from_inch(cls: Type[_TUnitMM], value: float) -> _TUnitMM:
        """
        Get instance from ``in`` (inch) value.

        Args:
            value (int): ``in`` value.

        Returns:
            UnitMM:
        """
        inst = super(UnitMM, cls).__new__(cls)  # type: ignore
        inst.__init__(UnitConvert.convert(num=value, frm=UnitLength.IN, to=UnitLength.MM))
        return inst

    @classmethod
    def from_inch10(cls: Type[_TUnitMM], value: float) -> _TUnitMM:
        """
        Get instance from ``1/10th in`` (inch) value.

        Args:
            value (int): ``1/10th in`` value.

        Returns:
            UnitMM:
        """
        inst = super(UnitMM, cls).__new__(cls)  # type: ignore
        inst.__init__(UnitConvert.convert(num=value, frm=UnitLength.IN10, to=UnitLength.MM))
        return inst

    @classmethod
    def from_inch100(cls: Type[_TUnitMM], value: float) -> _TUnitMM:
        """
        Get instance from ``1/100th in`` (inch) value.

        Args:
            value (int): ``1/100th in`` value.

        Returns:
            UnitMM:
        """
        inst = super(UnitMM, cls).__new__(cls)  # type: ignore
        inst.__init__(UnitConvert.convert(num=value, frm=UnitLength.IN100, to=UnitLength.MM))
        return inst

    @classmethod
    def from_inch1000(cls: Type[_TUnitMM], value: int) -> _TUnitMM:
        """
        Get instance from ``1/1,000th in`` (inch) value.

        Args:
            value (int): ``1/1,000th in`` value.

        Returns:
            UnitMM:
        """
        inst = super(UnitMM, cls).__new__(cls)  # type: ignore
        inst.__init__(UnitConvert.convert(num=value, frm=UnitLength.IN1000, to=UnitLength.MM))
        return inst

    @classmethod
    def from_cm(cls: Type[_TUnitMM], value: float) -> _TUnitMM:
        """
        Get instance from ``cm`` value.

        Args:
            value (int): ``cm`` value.

        Returns:
            UnitMM:
        """
        inst = super(UnitMM, cls).__new__(cls)  # type: ignore
        inst.__init__(UnitConvert.convert(num=value, frm=UnitLength.CM, to=UnitLength.MM))
        return inst

    @classmethod
    def from_unit_val(cls: Type[_TUnitMM], value: UnitT | float) -> _TUnitMM:
        """
        Get instance from ``UnitT`` or float value.

        Args:
            value (UnitT, float): ``UnitT`` or float value. If float then it is assumed to be in ``mm`` units.

        Returns:
            UnitMM:
        """
        try:
            unit_val = value.get_value_mm()  # type: ignore
            return cls.from_mm(unit_val)
        except AttributeError:
            return cls.from_mm(float(value))  # type: ignore
