from __future__ import annotations
import contextlib
from typing import TypeVar, Type, TYPE_CHECKING
from dataclasses import dataclass, field
from ooodev.utils.data_type.base_float_value import BaseFloatValue
from ooodev.units.unit_convert import UnitConvert
from ooodev.units.unit_convert import UnitLength

if TYPE_CHECKING:
    from ooodev.units.unit_obj import UnitT

_TUnitAppFont = TypeVar(name="_TUnitAppFont", bound="UnitAppFont")


@dataclass(unsafe_hash=True)
class UnitAppFont(BaseFloatValue):
    """
    Unit in ``AppFont`` units.

    Supports ``UnitT`` protocol.

    Warning:
        Although this class support ``UnitT`` protocol, ``get_unit_length()`` method returns ``UnitLength.INVALID``.

    Note:
        Unlike most other units in this module, this unit is not based on ``UnitLength``.
        This means that it does not have a valid ``UnitLength`` value and returns ``UnitLength.INVALID``.

        This unit require that the application font pixel ratio be set before it can be used.
        Which means office must be loaded before this unit can be used.

    See Also:
        :ref:`proto_unit_obj`
    """

    _ratio: float = field(init=False, repr=False, hash=False)

    def __post_init__(self):
        # Because most all other unit module do not need to access Lo, it is imported here.
        # This should allow other modules to be imported without needing Lo,
        # that is, if they don't call get_app_font() method.

        if not isinstance(self.value, float):
            object.__setattr__(self, "value", float(self.value))
        self._set_ratio()

    def _set_ratio(self) -> None:
        # pylint: disable=import-outside-toplevel
        from ooodev.loader.lo import Lo

        object.__setattr__(self, "_ratio", Lo.app_font_pixel_ratio)

    # region Overrides
    def _from_float(self, value: float) -> UnitAppFont:
        return self.from_app_font(value)

    # endregion Overrides

    # region math and comparison
    def __int__(self) -> int:
        return round(self.value)

    def __eq__(self, other: object) -> bool:
        if isinstance(other, UnitAppFont):
            return self.almost_equal(other.value)
        if hasattr(other, "get_value_app_font"):
            oth_val = other.get_value_app_font()  # type: ignore
            return self.almost_equal(oth_val)
        if hasattr(other, "get_value_mm100"):
            return self.get_value_mm100() == other.get_value_mm100()  # type: ignore
        with contextlib.suppress(Exception):
            return self.almost_equal(float(other))  # type: ignore
        return False

    def __add__(self, other: object) -> UnitAppFont:
        if isinstance(other, UnitAppFont):
            return self.from_app_font(self.value + other.value)
        if hasattr(other, "get_value_app_font"):
            oth_val = other.get_value_app_font()  # type: ignore
            return self.from_app_font(self.value + oth_val)
        if hasattr(other, "get_value_mm100"):
            oth_val = other.get_value_mm100()  # type: ignore
            oth_val_px = UnitConvert.convert(num=oth_val, frm=UnitLength.MM100, to=UnitLength.PX)
            return self.from_px(self.get_value_px() + oth_val_px)  # type: ignore

        if isinstance(other, (int, float)):
            return self.from_app_font(self.value + other)  # type: ignore
        return NotImplemented

    def __radd__(self, other: object) -> UnitAppFont:
        return self if other == 0 else self.__add__(other)

    def __sub__(self, other: object) -> UnitAppFont:
        if isinstance(other, UnitAppFont):
            return self.from_app_font(self.value - other.value)
        if hasattr(other, "get_value_app_font"):
            oth_val = other.get_value_app_font()  # type: ignore
            return self.from_app_font(self.value - oth_val)
        if hasattr(other, "get_value_mm100"):
            oth_val = other.get_value_mm100()  # type: ignore
            oth_val_px = UnitConvert.convert(num=oth_val, frm=UnitLength.MM100, to=UnitLength.PX)
            return self.from_px(self.get_value_px() - oth_val_px)  # type: ignore

        if isinstance(other, (int, float)):
            return self.from_app_font(self.value - other)  # type: ignore
        return NotImplemented

    def __rsub__(self, other: object) -> UnitAppFont:
        if isinstance(other, (int, float)):
            return self.from_app_font(other - self.value)  # type: ignore
        return NotImplemented

    def __mul__(self, other: object) -> UnitAppFont:
        if isinstance(other, UnitAppFont):
            return self.from_app_font(self.value * other.value)
        if hasattr(other, "get_value_app_font"):
            oth_val = other.get_value_app_font()  # type: ignore
            return self.from_app_font(self.value * oth_val)
        if hasattr(other, "get_value_mm100"):
            oth_val = other.get_value_mm100()  # type: ignore
            oth_val_px = UnitConvert.convert(num=oth_val, frm=UnitLength.MM100, to=UnitLength.PX)
            return self.from_px(self.get_value_px() * oth_val_px)

        if isinstance(other, (int, float)):
            return self.from_app_font(self.value * other)  # type: ignore

        return NotImplemented

    def __rmul__(self, other: int) -> UnitAppFont:
        return self if other == 0 else self.__mul__(other)

    def __truediv__(self, other: object) -> UnitAppFont:
        if isinstance(other, UnitAppFont):
            if other.value == 0:
                raise ZeroDivisionError
            return self.from_app_font(self.value / other.value)
        if hasattr(other, "get_value_app_font"):
            oth_val = other.get_value_app_font()  # type: ignore
            if oth_val == 0:
                raise ZeroDivisionError
            return self.from_app_font(self.value / oth_val)
        if hasattr(other, "get_value_mm100"):
            oth_val = other.get_value_mm100()  # type: ignore
            oth_val_px = UnitConvert.convert(num=oth_val, frm=UnitLength.MM100, to=UnitLength.PX)
            if oth_val_px == 0:
                raise ZeroDivisionError
            return self.from_mm(self.value / oth_val_px)  # type: ignore
        if isinstance(other, (int, float)):
            if other == 0:
                raise ZeroDivisionError
            return self.from_app_font(self.value / other)  # type: ignore
        return NotImplemented

    def __rtruediv__(self, other: object) -> UnitAppFont:
        if isinstance(other, (int, float)):
            if self.value == 0:
                raise ZeroDivisionError
            return self.from_app_font(other / self.value)  # type: ignore
        return NotImplemented

    def __abs__(self) -> float:
        return abs(self.value)

    def __lt__(self, other: object) -> bool:
        if isinstance(other, UnitAppFont):
            return self.value < other.value
        if hasattr(other, "get_value_app_font"):
            oth_val = other.get_value_app_font()  # type: ignore
            return self.value < oth_val
        if hasattr(other, "get_value_mm100"):
            return self.get_value_mm100() < other.get_value_mm100()  # type: ignore
        with contextlib.suppress(Exception):
            return self.value < float(other)  # type: ignore
        return False

    def __le__(self, other: object) -> bool:
        if isinstance(other, UnitAppFont):
            return True if self.almost_equal(other.value) else self.value < other.value
        if hasattr(other, "get_value_app_font"):
            oth_val = other.get_value_app_font()  # type: ignore
            return True if self.almost_equal(oth_val) else self.value < oth_val
        if hasattr(other, "get_value_mm100"):
            return self.get_value_mm100() <= other.get_value_mm100()  # type: ignore
        with contextlib.suppress(Exception):
            oth_val = float(other)  # type: ignore
            return True if self.almost_equal(oth_val) else self.value < oth_val
        return False

    def __gt__(self, other: object) -> bool:
        if isinstance(other, UnitAppFont):
            return self.value > other.value
        if hasattr(other, "get_value_app_font"):
            oth_val = other.get_value_app_font()  # type: ignore
            return self.value > oth_val
        if hasattr(other, "get_value_mm100"):
            return self.get_value_mm100() > other.get_value_mm100()  # type: ignore
        with contextlib.suppress(Exception):
            return self.value > float(other)  # type: ignore
        return False

    def __ge__(self, other: object) -> bool:
        if isinstance(other, UnitAppFont):
            return True if self.almost_equal(other.value) else self.value > other.value
        if hasattr(other, "get_value_app_font"):
            oth_val = other.get_value_app_font()  # type: ignore
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
            UnitLength: Instance unit length ``UnitLength.INVALID``.
        """
        return UnitLength.INVALID

    def convert_to(self, unit: UnitLength) -> float:
        """
        Converts instance value to specified unit.

        Args:
            unit (UnitLength): Unit to convert to.

        Returns:
            float: Value in specified unit.

        Hint:
            - ``UnitLength`` can be imported from ``ooodev.units``.
        """
        return UnitConvert.convert(num=self.get_value_px(), frm=UnitLength.PX, to=unit)

    def get_value_cm(self) -> float:
        """
        Gets instance value converted to ``cm`` units.

        Returns:
            float: Value in ``cm`` units.
        """
        return UnitConvert.convert(num=self.get_value_px(), frm=UnitLength.PX, to=UnitLength.CM)

    def get_value_mm100(self) -> int:
        """
        Gets instance value converted to Size in ``1/100th mm`` units.

        Returns:
            int: Value in ``1/100th mm`` units.
        """
        return round(UnitConvert.convert(num=self.get_value_px(), frm=UnitLength.PX, to=UnitLength.MM100))

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
        return UnitConvert.convert(num=self.get_value_px(), frm=UnitLength.PX, to=UnitLength.PT)

    def get_value_px(self) -> float:
        """
        Gets instance value in ``px`` (pixel) units.

        Returns:
            float: Value in ``px`` units.
        """
        return self.value * self._ratio

    def get_value_inch(self) -> float:
        """
        Gets instance value in ``in`` (inch) units.

        Returns:
            float: Value in ``in`` units.
        """
        return UnitConvert.convert(num=self.get_value_px(), frm=UnitLength.PX, to=UnitLength.IN)

    def get_value_inch100(self) -> int:
        """
        Gets instance value in ``1/100th inch`` units.

        Returns:
            int: Value in ``1/100th inch`` units.
        """
        return round(UnitConvert.convert(num=self.get_value_px(), frm=UnitLength.PX, to=UnitLength.IN100))

    def get_value_inch10(self) -> int:
        """
        Gets instance value in ``1/10th inch`` units.

        Returns:
            int: Value in ``1/10th inch`` units.
        """
        return round(UnitConvert.convert(num=self.get_value_px(), frm=UnitLength.PX, to=UnitLength.IN10))

    def get_value_inch1000(self) -> int:
        """
        Gets instance value in ``1/1000th inch`` units.

        Returns:
            int: Value in ``1/100th inch`` units.
        """
        return round(UnitConvert.convert(num=self.get_value_px(), frm=UnitLength.PX, to=UnitLength.IN1000))

    def get_value_app_font(self) -> float:
        """
        Gets instance value in ``AppFont`` units.

        Returns:
            float: Value in ``AppFont`` units.
        """
        return self.value

    @classmethod
    def from_mm(cls: Type[_TUnitAppFont], value: float) -> _TUnitAppFont:
        """
        Get instance from ``mm`` value.

        Args:
            value (int): ``mm`` value.

        Returns:
            UnitAppFont:
        """
        # pylint: disable=protected-access
        # pylint: disable=unnecessary-dunder-call
        px = UnitConvert.convert(num=value, frm=UnitLength.MM, to=UnitLength.PX)
        inst = super(UnitAppFont, cls).__new__(cls)  # type: ignore
        inst._set_ratio()
        app_font_val = px / inst._ratio
        inst.__init__(app_font_val)
        return inst

    @classmethod
    def from_mm10(cls: Type[_TUnitAppFont], value: float) -> _TUnitAppFont:
        """
        Get instance from ``1/10th mm`` value.

        Args:
            value (int): ``1/10th mm`` value.

        Returns:
            UnitAppFont:
        """
        # pylint: disable=protected-access
        # pylint: disable=unnecessary-dunder-call
        px = UnitConvert.convert(num=value, frm=UnitLength.MM10, to=UnitLength.PX)
        inst = super(UnitAppFont, cls).__new__(cls)  # type: ignore
        inst._set_ratio()
        app_font_val = px / inst._ratio
        inst.__init__(app_font_val)
        return inst

    @classmethod
    def from_mm100(cls: Type[_TUnitAppFont], value: int) -> _TUnitAppFont:
        """
        Get instance from ``1/100th mm`` value.

        Args:
            value (int): ``1/100th mm`` value.

        Returns:
            UnitAppFont:
        """
        # pylint: disable=protected-access
        # pylint: disable=unnecessary-dunder-call
        px = UnitConvert.convert(num=value, frm=UnitLength.MM100, to=UnitLength.PX)
        inst = super(UnitAppFont, cls).__new__(cls)  # type: ignore
        inst._set_ratio()
        app_font_val = px / inst._ratio
        inst.__init__(app_font_val)
        return inst

    @classmethod
    def from_pt(cls: Type[_TUnitAppFont], value: float) -> _TUnitAppFont:
        """
        Get instance from ``pt`` (points) value.

        Args:
            value (float): ``pt`` value.

        Returns:
            UnitAppFont:
        """
        # pylint: disable=protected-access
        # pylint: disable=unnecessary-dunder-call
        px = UnitConvert.convert(num=value, frm=UnitLength.PT, to=UnitLength.PX)
        inst = super(UnitAppFont, cls).__new__(cls)  # type: ignore
        inst._set_ratio()
        app_font_val = px / inst._ratio
        inst.__init__(app_font_val)
        return inst

    @classmethod
    def from_px(cls: Type[_TUnitAppFont], value: float) -> _TUnitAppFont:
        """
        Get instance from ``px`` (pixel) value.

        Args:
            value (float): ``px`` value.

        Returns:
            UnitAppFont:
        """
        # pylint: disable=protected-access
        # pylint: disable=unnecessary-dunder-call
        inst = super(UnitAppFont, cls).__new__(cls)  # type: ignore
        inst._set_ratio()
        app_font_val = value / inst._ratio
        inst.__init__(app_font_val)
        return inst

    @classmethod
    def from_inch(cls: Type[_TUnitAppFont], value: float) -> _TUnitAppFont:
        """
        Get instance from ``in`` (inch) value.

        Args:
            value (int): ``in`` value.

        Returns:
            UnitAppFont:
        """
        # pylint: disable=protected-access
        # pylint: disable=unnecessary-dunder-call
        px = UnitConvert.convert(num=value, frm=UnitLength.IN, to=UnitLength.PX)
        inst = super(UnitAppFont, cls).__new__(cls)  # type: ignore
        inst._set_ratio()
        app_font_val = px / inst._ratio
        inst.__init__(app_font_val)
        return inst

    @classmethod
    def from_inch10(cls: Type[_TUnitAppFont], value: float) -> _TUnitAppFont:
        """
        Get instance from ``1/10th in`` (inch) value.

        Args:
            value (int): ``1/10th in`` value.

        Returns:
            UnitAppFont:
        """
        # pylint: disable=protected-access
        # pylint: disable=unnecessary-dunder-call
        px = UnitConvert.convert(num=value, frm=UnitLength.IN10, to=UnitLength.PX)
        inst = super(UnitAppFont, cls).__new__(cls)  # type: ignore
        inst._set_ratio()
        app_font_val = px / inst._ratio
        inst.__init__(app_font_val)
        return inst

    @classmethod
    def from_inch100(cls: Type[_TUnitAppFont], value: float) -> _TUnitAppFont:
        """
        Get instance from ``1/100th in`` (inch) value.

        Args:
            value (int): ``1/100th in`` value.

        Returns:
            UnitAppFont:
        """
        # pylint: disable=protected-access
        # pylint: disable=unnecessary-dunder-call
        px = UnitConvert.convert(num=value, frm=UnitLength.IN100, to=UnitLength.PX)
        inst = super(UnitAppFont, cls).__new__(cls)  # type: ignore
        inst._set_ratio()
        app_font_val = px / inst._ratio
        inst.__init__(app_font_val)
        return inst

    @classmethod
    def from_inch1000(cls: Type[_TUnitAppFont], value: int) -> _TUnitAppFont:
        """
        Get instance from ``1/1,000th in`` (inch) value.

        Args:
            value (int): ``1/1,000th in`` value.

        Returns:
            UnitAppFont:
        """
        # pylint: disable=protected-access
        # pylint: disable=unnecessary-dunder-call
        px = UnitConvert.convert(num=value, frm=UnitLength.IN1000, to=UnitLength.PX)
        inst = super(UnitAppFont, cls).__new__(cls)  # type: ignore
        inst._set_ratio()
        app_font_val = px / inst._ratio
        inst.__init__(app_font_val)
        return inst

    @classmethod
    def from_cm(cls: Type[_TUnitAppFont], value: float) -> _TUnitAppFont:
        """
        Get instance from ``cm`` value.

        Args:
            value (int): ``cm`` value.

        Returns:
            UnitAppFont:
        """
        # pylint: disable=protected-access
        # pylint: disable=unnecessary-dunder-call
        px = UnitConvert.convert(num=value, frm=UnitLength.CM, to=UnitLength.PX)
        inst = super(UnitAppFont, cls).__new__(cls)  # type: ignore
        inst._set_ratio()
        app_font_val = px / inst._ratio
        inst.__init__(app_font_val)
        return inst

    @classmethod
    def from_app_font(cls: Type[_TUnitAppFont], value: float) -> _TUnitAppFont:
        """
        Get instance from ``AppFont`` value.

        Args:
            value (int): ``AppFont`` value.

        Returns:
            UnitAppFont:
        """
        # pylint: disable=unnecessary-dunder-call
        inst = super(UnitAppFont, cls).__new__(cls)  # type: ignore
        inst.__init__(value)
        return inst

    @classmethod
    def from_unit_val(cls: Type[_TUnitAppFont], value: UnitT | float | int) -> _TUnitAppFont:
        """
        Get instance from ``UnitT`` or float value.

        Args:
            value (UnitT, float, int): ``UnitT`` or float value. If float then it is assumed to be in ``mm`` units.

        Returns:
            UnitMM:
        """
        try:
            unit_val = value.get_value_px()  # type: ignore
            return cls.from_px(unit_val)
        except AttributeError:
            return cls.from_app_font(float(value))  # type: ignore
