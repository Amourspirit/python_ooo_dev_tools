from __future__ import annotations
import contextlib
from typing import TypeVar, Type, TYPE_CHECKING
from dataclasses import dataclass

try:
    # python 3.12+
    from typing import override  # noqa # type: ignore
except ImportError:
    from typing_extensions import override  # noqa # type: ignore

from ooodev.utils.data_type.base_float_value import BaseFloatValue
from ooodev.units.unit_convert import UnitConvert
from ooodev.units.unit_convert import UnitLength
from ooodev.utils.kind.point_size_kind import PointSizeKind

if TYPE_CHECKING:
    from ooodev.units.unit_obj import UnitT

_TUnitPX = TypeVar("_TUnitPX", bound="UnitPX")


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

    # region Overrides
    def _from_float(self, value: float) -> UnitPX:
        return self.from_px(value)

    # endregion Overrides

    # region math and comparison
    @override
    def __int__(self) -> int:
        return round(self.value)

    @override
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

    @override
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

    @override
    def __radd__(self, other: object) -> UnitPX:
        return self if other == 0 else self.__add__(other)

    @override
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

    @override
    def __rsub__(self, other: object) -> UnitPX:
        if isinstance(other, (int, float)):
            return self.from_px(other - self.value)  # type: ignore
        return NotImplemented

    @override
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

    @override
    def __rmul__(self, other: int) -> UnitPX:
        return self if other == 0 else self.__mul__(other)

    @override
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

    @override
    def __rtruediv__(self, other: object) -> UnitPX:
        if isinstance(other, (int, float)):
            if self.value == 0:
                raise ZeroDivisionError
            return self.from_px(other / self.value)  # type: ignore
        return NotImplemented

    @override
    def __abs__(self) -> float:  # type: ignore
        return abs(self.value)

    @override
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

    @override
    def __le__(self, other: object) -> bool:
        if isinstance(other, UnitPX):
            return True if self.almost_equal(other.value) else self.value < other.value
        if hasattr(other, "get_value_px"):
            oth_val = other.get_value_px()  # type: ignore
            return True if self.almost_equal(oth_val) else self.value < oth_val
        if hasattr(other, "get_value_mm100"):
            return self.get_value_mm100() <= other.get_value_mm100()  # type: ignore
        with contextlib.suppress(Exception):
            oth_val = float(other)  # type: ignore
            return True if self.almost_equal(oth_val) else self.value < oth_val
        return False

    @override
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

    @override
    def __ge__(self, other: object) -> bool:
        if isinstance(other, UnitPX):
            return True if self.almost_equal(other.value) else self.value > other.value
        if hasattr(other, "get_value_px"):
            oth_val = other.get_value_px()  # type: ignore
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
            UnitLength: Instance unit length ``UnitLength.PX``.
        """
        return UnitLength.PX

    def convert_to(self, unit: UnitLength) -> float:
        """
        Converts instance value to specified unit.

        Args:
            unit (UnitLength): Unit to convert to.

        Returns:
            float: Value in specified unit.
        """
        return UnitConvert.convert(num=self.value, frm=UnitLength.PX, to=unit)

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

    def get_value_app_font(self, kind: PointSizeKind | int) -> float:
        """
        Gets instance value in ``AppFont`` units.

        Returns:
            float: Value in ``AppFont`` units.
            kind (PointSizeKind, optional): The kind of ``AppFont`` to use.

        Note:
            AppFont units have different values when converted.
            This is true even if they have the same value in ``AppFont`` units.
            ``AppFontX(10)`` is not equal to ``AppFontY(10)`` when they are converted to different units.

            ``Kind`` when ``int`` is used, the value must be one of the following:

            - ``0`` is ``PointSizeKind.X``,
            - ``1`` is ``PointSizeKind.Y``,
            - ``2`` is ``PointSizeKind.WIDTH``,
            - ``3`` is ``PointSizeKind.HEIGHT``.

        Hint:
            - ``PointSizeKind`` can be imported from ``ooodev.utils.kind.point_size_kind``.
        """
        # pylint: disable=import-outside-toplevel
        from ooodev.units._app_font.app_font_factory import AppFontFactory

        af_type = AppFontFactory.get_app_font_type(kind=kind)
        af = af_type.from_px(self.value)
        return af.value

    @classmethod
    def from_pt(cls: Type[_TUnitPX], value: float) -> _TUnitPX:
        """
        Get instance from ``pt`` (points) value.

        Args:
            value (float): ``pt`` value.

        Returns:
            UnitPX:
        """
        # pylint: disable=unnecessary-dunder-call
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
        # pylint: disable=unnecessary-dunder-call
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
        # pylint: disable=unnecessary-dunder-call
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
        # pylint: disable=unnecessary-dunder-call
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
        # pylint: disable=unnecessary-dunder-call
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
        # pylint: disable=unnecessary-dunder-call
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
        # pylint: disable=unnecessary-dunder-call
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
        # pylint: disable=unnecessary-dunder-call
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
        # pylint: disable=unnecessary-dunder-call
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
        # pylint: disable=unnecessary-dunder-call
        inst = super(UnitPX, cls).__new__(cls)  # type: ignore
        inst.__init__(UnitConvert.convert(num=value, frm=UnitLength.CM, to=UnitLength.PX))
        return inst

    @classmethod
    def from_app_font(cls: Type[_TUnitPX], value: float, kind: PointSizeKind | int) -> _TUnitPX:
        """
        Get instance from ``AppFont`` value.

        Args:
            value (int): ``AppFont`` value.
            kind (PointSizeKind, optional): The kind of ``AppFont`` to use.

        Returns:
            UnitPX:

        Note:
            AppFont units have different values when converted.
            This is true even if they have the same value in ``AppFont`` units.
            ``AppFontX(10)`` is not equal to ``AppFontY(10)`` when they are converted to different units.

            ``Kind`` when ``int`` is used, the value must be one of the following:

            - ``0`` is ``PointSizeKind.X``,
            - ``1`` is ``PointSizeKind.Y``,
            - ``2`` is ``PointSizeKind.WIDTH``,
            - ``3`` is ``PointSizeKind.HEIGHT``.

        Hint:
            - ``PointSizeKind`` can be imported from ``ooodev.utils.kind.point_size_kind``.
        """
        # pylint: disable=import-outside-toplevel
        from ooodev.units._app_font.app_font_factory import AppFontFactory

        af = AppFontFactory.get_app_font(kind=kind, val=value)
        return cls.from_px(af.get_value_px())

    @classmethod
    def from_unit_val(cls: Type[_TUnitPX], value: UnitT | float | int) -> _TUnitPX:
        """
        Get instance from ``UnitT`` or float value.

        Args:
            value (UnitT, float | int): ``UnitT`` or float value. If float then it is assumed to be in ``px`` units.

        Returns:
            UnitPX:
        """
        try:
            unit_val = value.get_value_px()  # type: ignore
            return cls.from_px(unit_val)
        except AttributeError:
            return cls.from_px(float(value))  # type: ignore
