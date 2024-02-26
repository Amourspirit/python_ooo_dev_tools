from __future__ import annotations
import contextlib
from typing import TypeVar, Type, TYPE_CHECKING
from dataclasses import dataclass
from ooodev.utils.data_type.base_float_value import BaseFloatValue
from ooodev.units.unit_convert import UnitConvert
from ooodev.units.unit_convert import UnitLength

if TYPE_CHECKING:
    from ooodev.units.unit_obj import UnitT

_TUnitPX = TypeVar(name="_TUnitPX", bound="UnitPX")


@dataclass(unsafe_hash=True)
class UnitPX(BaseFloatValue):
    """
    Represents a ``PX`` (pixel) value.

    Supports ``UnitT`` protocol.

    See Also:
        :ref:`proto_unit_obj`
    """

    def __post_init__(self) -> None:
        if not isinstance(self.value, float):
            object.__setattr__(self, "value", float(self.value))

    # region math and comparison
    def __int__(self) -> int:
        return round(self.value)

    def __eq__(self, other: object) -> bool:
        if isinstance(other, UnitPX):
            return self.almost_equal(other.value)
        if hasattr(other, "get_value_px"):
            oth_val = other.get_value_px()  # type: ignore
            return self.almost_equal(oth_val)
        if hasattr(other, "get_value_mm100"):
            return self.get_value_mm100() == other.get_value_mm100()  # type: ignore
        with contextlib.suppress(Exception):
            return self.almost_equal(float(other))  # type: ignore
        return False

    def __add__(self, other: object) -> UnitPX:
        if isinstance(other, UnitPX):
            return self.from_px(self.value + other.value)
        if hasattr(other, "get_value_px"):
            oth_val = other.get_value_px()  # type: ignore
            return self.from_px(self.value + oth_val)
        if hasattr(other, "get_value_mm100"):
            oth_val = other.get_value_mm100()  # type: ignore
            oth_val_px = UnitConvert.convert(num=oth_val, frm=UnitLength.MM100, to=UnitLength.PX)
            return self.from_px(self.value + oth_val_px)  # type: ignore

        if isinstance(other, (int, float)):
            return self.from_px(self.value + other)  # type: ignore
        return NotImplemented

    def __radd__(self, other: object) -> UnitPX:
        return self if other == 0 else self.__add__(other)

    def __sub__(self, other: object) -> UnitPX:
        if isinstance(other, UnitPX):
            return self.from_px(self.value - other.value)
        if hasattr(other, "get_value_px"):
            oth_val = other.get_value_px()  # type: ignore
            return self.from_px(self.value - oth_val)
        if hasattr(other, "get_value_mm100"):
            oth_val = other.get_value_mm100()  # type: ignore
            oth_val_px = UnitConvert.convert(num=oth_val, frm=UnitLength.MM100, to=UnitLength.PX)
            return self.from_px(self.value - oth_val_px)  # type: ignore

        if isinstance(other, (int, float)):
            return self.from_px(self.value - other)  # type: ignore
        return NotImplemented

    def __rsub__(self, other: object) -> UnitPX:
        if isinstance(other, (int, float)):
            return self.from_px(other - self.value)  # type: ignore
        return NotImplemented

    def __mul__(self, other: object) -> UnitPX:
        if isinstance(other, UnitPX):
            return self.from_px(self.value * other.value)
        if hasattr(other, "get_value_px"):
            oth_val = other.get_value_px()  # type: ignore
            return self.from_px(self.value * oth_val)
        if hasattr(other, "get_value_mm100"):
            oth_val = other.get_value_mm100()  # type: ignore
            oth_val_px = UnitConvert.convert(num=oth_val, frm=UnitLength.MM100, to=UnitLength.PX)
            return self.from_px(self.value * oth_val_px)

        if isinstance(other, (int, float)):
            return self.from_px(self.value * other)  # type: ignore

        return NotImplemented

    def __rmul__(self, other: int) -> UnitPX:
        return self if other == 0 else self.__mul__(other)

    def __truediv__(self, other: object) -> UnitPX:
        if isinstance(other, UnitPX):
            if other.value == 0:
                raise ZeroDivisionError
            return self.from_px(self.value / other.value)
        if hasattr(other, "get_value_px"):
            oth_val = other.get_value_px()  # type: ignore
            if oth_val == 0:
                raise ZeroDivisionError
            return self.from_px(self.value / oth_val)
        if hasattr(other, "get_value_mm100"):
            oth_val = other.get_value_mm100()  # type: ignore
            oth_val_px = UnitConvert.convert(num=oth_val, frm=UnitLength.MM100, to=UnitLength.PX)
            if oth_val_px == 0:
                raise ZeroDivisionError
            return self.from_px(self.value / oth_val_px)  # type: ignore
        if isinstance(other, (int, float)):
            if other == 0:
                raise ZeroDivisionError
            return self.from_px(self.value / other)  # type: ignore
        return NotImplemented

    def __rtruediv__(self, other: object) -> UnitPX:
        if isinstance(other, (int, float)):
            if self.value == 0:
                raise ZeroDivisionError
            return self.from_px(other / self.value)  # type: ignore
        return NotImplemented

    def __abs__(self) -> float:
        return abs(self.value)

    def __lt__(self, other: object) -> bool:
        if isinstance(other, UnitPX):
            return self.value < other.value
        if hasattr(other, "get_value_px"):
            oth_val = other.get_value_px()  # type: ignore
            return self.value < oth_val
        if hasattr(other, "get_value_mm100"):
            return self.get_value_mm100() < other.get_value_mm100()  # type: ignore
        with contextlib.suppress(Exception):
            return self.value < float(other)  # type: ignore
        return False

    def __le__(self, other: object) -> bool:
        if isinstance(other, UnitPX):
            if self.almost_equal(other.value):
                return True
            return self.value < other.value
        if hasattr(other, "get_value_px"):
            oth_val = other.get_value_px()  # type: ignore
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
        if isinstance(other, UnitPX):
            return self.value > other.value
        if hasattr(other, "get_value_px"):
            oth_val = other.get_value_px()  # type: ignore
            return self.value > oth_val
        if hasattr(other, "get_value_mm100"):
            return self.get_value_mm100() > other.get_value_mm100()  # type: ignore
        with contextlib.suppress(Exception):
            return self.value > float(other)  # type: ignore
        return False

    def __ge__(self, other: object) -> bool:
        if isinstance(other, UnitPX):
            if self.almost_equal(other.value):
                return True
            return self.value > other.value
        if hasattr(other, "get_value_px"):
            oth_val = other.get_value_px()  # type: ignore
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
        return UnitConvert.convert(num=self.value, frm=UnitLength.PX, to=UnitLength.CM)

    def get_value_mm(self) -> float:
        """
        Gets instance value converted to ``mm`` units.

        Returns:
            int: Value in ``mm`` units.
        """
        return float(UnitConvert.convert(num=self.value, frm=UnitLength.PX, to=UnitLength.MM))

    def get_value_mm100(self) -> int:
        """
        Gets instance value converted to ``1/100th mm`` units.

        Returns:
            int: Value in ``1/100th mm`` units.
        """
        return round(UnitConvert.convert(num=self.value, frm=UnitLength.PX, to=UnitLength.MM100))

    def get_value_pt(self) -> float:
        """
        Gets instance value in ``pt`` (point) units.

        Returns:
            int: Value in ``pt`` units.
        """
        return UnitConvert.convert(num=self.value, frm=UnitLength.PX, to=UnitLength.PT)

    def get_value_px(self) -> float:
        """
        Gets instance value in ``px`` (pixel) units.

        Returns:
            int: Value in ``px`` units.
        """
        return self.value

    def get_value_inch(self) -> float:
        """
        Gets instance value in ``in`` (inch) units.

        Returns:
            float: Value in ``in`` units.
        """
        return UnitConvert.convert(num=self.value, frm=UnitLength.PX, to=UnitLength.IN)

    def get_value_inch100(self) -> int:
        """
        Gets instance value in ``1/100th inch`` units.

        Returns:
            int: Value in ``1/100th inch`` units.
        """
        return round(UnitConvert.convert(num=self.value, frm=UnitLength.PX, to=UnitLength.IN100))

    def get_value_inch10(self) -> int:
        """
        Gets instance value in ``1/10th inch`` units.

        Returns:
            int: Value in ``1/10th inch`` units.
        """
        return round(UnitConvert.convert(num=self.value, frm=UnitLength.PX, to=UnitLength.IN10))

    def get_value_inch1000(self) -> int:
        """
        Gets instance value in ``1/1000th inch`` units.

        Returns:
            int: Value in ``1/100th inch`` units.
        """
        return round(UnitConvert.convert(num=self.value, frm=UnitLength.PX, to=UnitLength.IN1000))

    @classmethod
    def from_pt(cls: Type[_TUnitPX], value: float) -> _TUnitPX:
        """
        Get instance from ``pt`` (points) value.

        Args:
            value (float): ``pt`` value.

        Returns:
            UnitPX:
        """
        inst = super(UnitPX, cls).__new__(cls)  # type: ignore
        inst.__init__(UnitConvert.convert(num=value, frm=UnitLength.PT, to=UnitLength.PX))
        return inst

    @classmethod
    def from_px(cls: Type[_TUnitPX], value: float) -> _TUnitPX:
        """
        Get instance from ``px`` (pixel) value.

        Args:
            value (float): ``px`` value.

        Returns:
            UnitPX:
        """
        inst = super(UnitPX, cls).__new__(cls)  # type: ignore
        inst.__init__(value)
        return inst

    @classmethod
    def from_mm(cls: Type[_TUnitPX], value: float) -> _TUnitPX:
        """
        Get instance from ``mm`` value.

        Args:
            value (int): ``mm`` value.

        Returns:
            UnitPX:
        """
        inst = super(UnitPX, cls).__new__(cls)  # type: ignore
        inst.__init__(UnitConvert.convert(num=value, frm=UnitLength.MM, to=UnitLength.PX))
        return inst

    @classmethod
    def from_mm10(cls: Type[_TUnitPX], value: float) -> _TUnitPX:
        """
        Get instance from ``1/10th mm`` value.

        Args:
            value (int): ``1/10th mm`` value.

        Returns:
            UnitPX:
        """
        inst = super(UnitPX, cls).__new__(cls)  # type: ignore
        inst.__init__(UnitConvert.convert(num=value, frm=UnitLength.MM10, to=UnitLength.PX))
        return inst

    @classmethod
    def from_mm100(cls: Type[_TUnitPX], value: int) -> _TUnitPX:
        """
        Get instance from ``1/100th mm`` value.

        Args:
            value (int): ``1/100th mm`` value.

        Returns:
            UnitPX:
        """
        inst = super(UnitPX, cls).__new__(cls)  # type: ignore
        inst.__init__(UnitConvert.convert(num=value, frm=UnitLength.MM100, to=UnitLength.PX))
        return inst

    @classmethod
    def from_inch(cls: Type[_TUnitPX], value: float) -> _TUnitPX:
        """
        Get instance from ``in`` (inch) value.

        Args:
            value (int): ``in`` value.

        Returns:
            UnitPX:
        """
        inst = super(UnitPX, cls).__new__(cls)  # type: ignore
        inst.__init__(UnitConvert.convert(num=value, frm=UnitLength.IN, to=UnitLength.PX))
        return inst

    @classmethod
    def from_inch10(cls: Type[_TUnitPX], value: float) -> _TUnitPX:
        """
        Get instance from ``1/10th in`` (inch) value.

        Args:
            value (int): ``1/10th in`` value.

        Returns:
            UnitPX:
        """
        inst = super(UnitPX, cls).__new__(cls)  # type: ignore
        inst.__init__(UnitConvert.convert(num=value, frm=UnitLength.IN10, to=UnitLength.PX))
        return inst

    @classmethod
    def from_inch100(cls: Type[_TUnitPX], value: float) -> _TUnitPX:
        """
        Get instance from ``1/100th in`` (inch) value.

        Args:
            value (int): ``1/100th in`` value.

        Returns:
            UnitPX:
        """
        inst = super(UnitPX, cls).__new__(cls)  # type: ignore
        inst.__init__(UnitConvert.convert(num=value, frm=UnitLength.IN100, to=UnitLength.PX))
        return inst

    @classmethod
    def from_inch1000(cls: Type[_TUnitPX], value: int) -> _TUnitPX:
        """
        Get instance from ``1/1,000th in`` (inch) value.

        Args:
            value (int): ``1/1,000th in`` value.

        Returns:
            UnitPX:
        """
        inst = super(UnitPX, cls).__new__(cls)  # type: ignore
        inst.__init__(UnitConvert.convert(num=value, frm=UnitLength.IN1000, to=UnitLength.PX))
        return inst

    @classmethod
    def from_cm(cls: Type[_TUnitPX], value: float) -> _TUnitPX:
        """
        Get instance from ``cm`` value.

        Args:
            value (int): ``cm`` value.

        Returns:
            UnitPX:
        """
        inst = super(UnitPX, cls).__new__(cls)  # type: ignore
        inst.__init__(UnitConvert.convert(num=value, frm=UnitLength.CM, to=UnitLength.PX))
        return inst

    @classmethod
    def from_unit_val(cls: Type[_TUnitPX], value: UnitT | float) -> _TUnitPX:
        """
        Get instance from ``UnitT`` or float value.

        Args:
            value (UnitT, float): ``UnitT`` or float value. If float then it is assumed to be in ``px`` units.

        Returns:
            UnitPX:
        """
        try:
            unit_val = value.get_value_px()  # type: ignore
            return cls.from_px(unit_val)
        except AttributeError:
            return cls.from_px(float(value))  # type: ignore
