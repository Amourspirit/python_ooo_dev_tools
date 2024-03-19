from __future__ import annotations
import contextlib
from typing import TypeVar, Type, TYPE_CHECKING
from dataclasses import dataclass
from ooodev.utils.data_type.base_float_value import BaseFloatValue
from ooodev.units.unit_convert import UnitConvert
from ooodev.units.unit_convert import UnitLength


if TYPE_CHECKING:
    from ooodev.units.unit_obj import UnitT

_TUnitInch = TypeVar(name="_TUnitInch", bound="UnitInch")


@dataclass(unsafe_hash=True)
class UnitInch(BaseFloatValue):
    """
    Unit in ``inch`` units.

    Supports ``UnitT`` protocol.

    See Also:
        :ref:`proto_unit_obj`
    """

    def __post_init__(self):
        if not isinstance(self.value, float):
            object.__setattr__(self, "value", float(self.value))

    # region Overrides
    def _from_float(self, value: float) -> UnitInch:
        return self.from_inch(value)

    # region math and comparison
    def __int__(self) -> int:
        return round(self.value)

    def __eq__(self, other: object) -> bool:
        if isinstance(other, UnitInch):
            return self.almost_equal(other.value)
        if hasattr(other, "get_value_inch"):
            oth_val = other.get_value_inch()  # type: ignore
            return self.almost_equal(oth_val)
        if hasattr(other, "get_value_mm100"):
            return self.get_value_mm100() == other.get_value_mm100()  # type: ignore
        with contextlib.suppress(Exception):
            return self.almost_equal(float(other))  # type: ignore
        return False

    def __add__(self, other: object) -> UnitInch:
        if isinstance(other, UnitInch):
            return self.from_inch(self.value + other.value)
        if hasattr(other, "get_value_inch"):
            oth_val = other.get_value_inch()  # type: ignore
            return self.from_inch(self.value + oth_val)
        if hasattr(other, "get_value_mm100"):
            oth_val = other.get_value_mm100()  # type: ignore
            oth_val_inch = UnitConvert.convert(num=oth_val, frm=UnitLength.MM100, to=UnitLength.IN)
            return self.from_inch(self.value + oth_val_inch)  # type: ignore

        if isinstance(other, (int, float)):
            return self.from_inch(self.value + other)  # type: ignore
        return NotImplemented

    def __radd__(self, other: object) -> UnitInch:
        return self if other == 0 else self.__add__(other)

    def __sub__(self, other: object) -> UnitInch:
        if isinstance(other, UnitInch):
            return self.from_inch(self.value - other.value)
        if hasattr(other, "get_value_inch"):
            oth_val = other.get_value_inch()  # type: ignore
            return self.from_inch(self.value - oth_val)
        if hasattr(other, "get_value_mm100"):
            oth_val = other.get_value_mm100()  # type: ignore
            oth_val_inch = UnitConvert.convert(num=oth_val, frm=UnitLength.MM100, to=UnitLength.IN)
            return self.from_inch(self.value - oth_val_inch)  # type: ignore

        if isinstance(other, (int, float)):
            return self.from_inch(self.value - other)  # type: ignore
        return NotImplemented

    def __rsub__(self, other: object) -> UnitInch:
        if isinstance(other, (int, float)):
            return self.from_inch(other - self.value)  # type: ignore
        return NotImplemented

    def __mul__(self, other: object) -> UnitInch:
        if isinstance(other, UnitInch):
            return self.from_inch(self.value * other.value)
        if hasattr(other, "get_value_inch"):
            oth_val = other.get_value_inch()  # type: ignore
            return self.from_inch(self.value * oth_val)
        if hasattr(other, "get_value_mm100"):
            oth_val = other.get_value_mm100()  # type: ignore
            oth_val_inch = UnitConvert.convert(num=oth_val, frm=UnitLength.MM100, to=UnitLength.IN)
            return self.from_inch(self.value * oth_val_inch)  # type: ignore

        if isinstance(other, (int, float)):
            return self.from_inch(self.value * other)  # type: ignore

        return NotImplemented

    def __rmul__(self, other: int) -> UnitInch:
        return self if other == 0 else self.__mul__(other)

    def __truediv__(self, other: object) -> UnitInch:
        if isinstance(other, UnitInch):
            if other.value == 0:
                raise ZeroDivisionError
            return self.from_inch(self.value / other.value)
        if hasattr(other, "get_value_inch"):
            oth_val = other.get_value_inch()  # type: ignore
            if oth_val == 0:
                raise ZeroDivisionError
            return self.from_inch(self.value / oth_val)
        if hasattr(other, "get_value_mm100"):
            oth_val = other.get_value_mm100()  # type: ignore
            oth_val_inch = UnitConvert.convert(num=oth_val, frm=UnitLength.MM100, to=UnitLength.IN)
            if oth_val_inch == 0:
                raise ZeroDivisionError
            return self.from_inch(self.value / oth_val_inch)  # type: ignore
        if isinstance(other, (int, float)):
            if other == 0:
                raise ZeroDivisionError
            return self.from_inch(self.value / other)  # type: ignore
        return NotImplemented

    def __rtruediv__(self, other: object) -> UnitInch:
        if isinstance(other, (int, float)):
            if self.value == 0:
                raise ZeroDivisionError
            return self.from_inch(other / self.value)  # type: ignore
        return NotImplemented

    def __abs__(self) -> float:
        return abs(self.value)

    def __lt__(self, other: object) -> bool:
        if isinstance(other, UnitInch):
            return self.value < other.value
        if hasattr(other, "get_value_inch"):
            oth_val = other.get_value_inch()  # type: ignore
            return self.value < oth_val
        if hasattr(other, "get_value_mm100"):
            return self.get_value_mm100() < other.get_value_mm100()  # type: ignore
        with contextlib.suppress(Exception):
            return self.value < float(other)  # type: ignore
        return False

    def __le__(self, other: object) -> bool:
        if isinstance(other, UnitInch):
            return True if self.almost_equal(other.value) else self.value < other.value
        if hasattr(other, "get_value_inch"):
            oth_val = other.get_value_inch()  # type: ignore
            return True if self.almost_equal(oth_val) else self.value < oth_val
        if hasattr(other, "get_value_mm100"):
            return self.get_value_mm100() <= other.get_value_mm100()  # type: ignore
        with contextlib.suppress(Exception):
            oth_val = float(other)  # type: ignore
            return True if self.almost_equal(oth_val) else self.value < oth_val
        return False

    def __gt__(self, other: object) -> bool:
        if isinstance(other, UnitInch):
            return self.value > other.value
        if hasattr(other, "get_value_inch"):
            oth_val = other.get_value_inch()  # type: ignore
            return self.value > oth_val
        if hasattr(other, "get_value_mm100"):
            return self.get_value_mm100() > other.get_value_mm100()  # type: ignore
        with contextlib.suppress(Exception):
            return self.value > float(other)  # type: ignore
        return False

    def __ge__(self, other: object) -> bool:
        if isinstance(other, UnitInch):
            return True if self.almost_equal(other.value) else self.value > other.value
        if hasattr(other, "get_value_inch"):
            oth_val = other.get_value_inch()  # type: ignore
            return True if self.almost_equal(oth_val) else self.value > oth_val
        if hasattr(other, "get_value_mm100"):
            return self.get_value_mm100() >= other.get_value_mm100()  # type: ignore
        with contextlib.suppress(Exception):
            oth_val = float(other)  # type: ignore
            return True if self.almost_equal(oth_val) else self.value > oth_val
        return False

    # endregion math and comparison

    @staticmethod
    def get_unit_length() -> UnitLength:
        """
        Gets instance unit length.

        Returns:
            UnitLength: Instance unit length ``UnitLength.IN``.
        """
        return UnitLength.IN

    def convert_to(self, unit: UnitLength) -> float:
        """
        Converts instance value to specified unit.

        Args:
            unit (UnitLength): Unit to convert to.

        Returns:
            float: Value in specified unit.
        """
        return UnitConvert.convert(num=self.value, frm=UnitLength.IN, to=unit)

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

    def get_value_inch(self) -> float:
        """
        Gets instance value in inch units.

        Returns:
            float: Value in inch units.
        """
        return self.value

    def get_value_inch10(self) -> float:
        """
        Gets instance value in ``1/10th inch`` units.

        Returns:
            float: Value in ``1/10th inch`` units.
        """
        return UnitConvert.convert(num=self.value, frm=UnitLength.IN, to=UnitLength.IN10)

    def get_value_inch100(self) -> float:
        """
        Gets instance value in ``1/100th inch`` units.

        Returns:
            float: Value in ``1/100th inch`` units.
        """
        return UnitConvert.convert(num=self.value, frm=UnitLength.IN, to=UnitLength.IN100)

    def get_value_inch1000(self) -> int:
        """
        Gets instance value in ``1/1000th inch`` units.

        Returns:
            int: Value in ``1/100th inch`` units.
        """
        return round(UnitConvert.convert(num=self.value, frm=UnitLength.IN, to=UnitLength.IN1000))

    def get_value_app_font(self, x: bool = True) -> float:
        """
        Gets instance value in ``AppFont`` units.

        Returns:
            float: Value in ``AppFont`` units.
            x (bool, optional): If ``True`` then ``AppFontX`` is used else ``AppFontY`` is used. Defaults to ``True``.

        Note:
            ``AppFontX`` and ``AppFontY`` have different values when converted.
            This is true even if they have the same value in ``AppFont`` units.
            ``AppFontX(10)`` is not equal to ``AppFontY(10)`` when they are converted to different units.
        """
        # pylint: disable=import-outside-toplevel
        if x:
            from ooodev.units.unit_app_font_x import UnitAppFontX

            af = UnitAppFontX.from_inch(self.value)
        else:
            from ooodev.units.unit_app_font_y import UnitAppFontY

            af = UnitAppFontY.from_inch(self.value)
        return af.value

    @classmethod
    def from_mm(cls: Type[_TUnitInch], value: float) -> _TUnitInch:
        """
        Get instance from ``mm`` value.

        Args:
            value (float): ``mm`` value.

        Returns:
            UnitInch:
        """
        # pylint: disable=unnecessary-dunder-call
        inst = super(UnitInch, cls).__new__(cls)  # type: ignore
        inst.__init__(UnitConvert.convert(num=value, frm=UnitLength.MM, to=UnitLength.IN))
        return inst

    @classmethod
    def from_mm10(cls: Type[_TUnitInch], value: int) -> _TUnitInch:
        """
        Get instance from ``1/10th mm`` value.

        Args:
            value (int): ``1/10th mm`` value.

        Returns:
            UnitInch:
        """
        # pylint: disable=unnecessary-dunder-call
        inst = super(UnitInch, cls).__new__(cls)  # type: ignore
        inst.__init__(UnitConvert.convert(num=value, frm=UnitLength.MM10, to=UnitLength.IN))
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
        # pylint: disable=unnecessary-dunder-call
        inst = super(UnitInch, cls).__new__(cls)  # type: ignore
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
        # pylint: disable=unnecessary-dunder-call
        inst = super(UnitInch, cls).__new__(cls)  # type: ignore
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
        # pylint: disable=unnecessary-dunder-call
        inst = super(UnitInch, cls).__new__(cls)  # type: ignore
        inst.__init__(UnitConvert.convert(num=value, frm=UnitLength.PX, to=UnitLength.IN))
        return inst

    @classmethod
    def from_inch(cls: Type[_TUnitInch], value: float) -> _TUnitInch:
        """
        Get instance from inch value.

        Args:
            value (int): Inch value.

        Returns:
            UnitInch:
        """
        # pylint: disable=unnecessary-dunder-call
        inst = super(UnitInch, cls).__new__(cls)  # type: ignore
        inst.__init__(value)
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
        # pylint: disable=unnecessary-dunder-call
        inst = super(UnitInch, cls).__new__(cls)  # type: ignore
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
        # pylint: disable=unnecessary-dunder-call
        inst = super(UnitInch, cls).__new__(cls)  # type: ignore
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
        # pylint: disable=unnecessary-dunder-call
        inst = super(UnitInch, cls).__new__(cls)  # type: ignore
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
        # pylint: disable=unnecessary-dunder-call
        inst = super(UnitInch, cls).__new__(cls)  # type: ignore
        inst.__init__(UnitConvert.convert(num=value, frm=UnitLength.CM, to=UnitLength.IN))
        return inst

    @classmethod
    def from_app_font(cls: Type[_TUnitInch], value: float, x: bool = True) -> _TUnitInch:
        """
        Get instance from ``AppFont`` value.

        Args:
            value (int): ``AppFont`` value.
            x (bool, optional): If ``True`` then ``AppFontX`` is used else ``AppFontY`` is used. Defaults to ``True``.

        Returns:
            UnitInch:

        Note:
            ``AppFontX`` and ``AppFontY`` have different values when converted.
            This is true even if they have the same value in ``AppFont`` units.
            ``AppFontX(10)`` is not equal to ``AppFontY(10)`` when they are converted to different units.
        """
        # pylint: disable=import-outside-toplevel

        if x:
            from ooodev.units.unit_app_font_x import UnitAppFontX

            af = UnitAppFontX(value)
        else:
            from ooodev.units.unit_app_font_y import UnitAppFontY

            af = UnitAppFontY(value)
        return cls.from_inch(af.get_value_inch())

    @classmethod
    def from_unit_val(cls: Type[_TUnitInch], value: UnitT | float | int) -> _TUnitInch:
        """
        Get instance from ``UnitT`` or float value.

        Args:
            value (UnitT, float, int): ``UnitT`` or float value. If float then it is assumed to be in ``inch`` units.

        Returns:
            UnitInch:
        """
        try:
            if hasattr(value, "get_value_inch"):
                return cls.from_inch(value.get_value_inch())  # type: ignore

            unit_100 = value.get_value_mm100()  # type: ignore
            return cls.from_mm100(unit_100)
        except AttributeError:
            return cls.from_inch(float(value))  # type: ignore
