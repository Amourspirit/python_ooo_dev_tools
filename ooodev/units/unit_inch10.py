from __future__ import annotations
import contextlib
from typing import TypeVar, Type, TYPE_CHECKING
from dataclasses import dataclass
from ooodev.utils.data_type.base_float_value import BaseFloatValue
from ooodev.units.unit_convert import UnitConvert
from ooodev.units.unit_convert import UnitLength
from ooodev.utils.kind.point_size_kind import PointSizeKind

if TYPE_CHECKING:
    from ooodev.units.unit_obj import UnitT

_TUnitInch10 = TypeVar(name="_TUnitInch10", bound="UnitInch10")


@dataclass(unsafe_hash=True)
class UnitInch10(BaseFloatValue):
    """
    Unit in ``1/10th in`` units.

    Supports ``UnitT`` protocol.

    See Also:
        :ref:`proto_unit_obj`
    """

    def __post_init__(self):
        if not isinstance(self.value, float):
            object.__setattr__(self, "value", float(self.value))

    # region Overrides
    def _from_float(self, value: float) -> UnitInch10:
        return self.from_inch10(value)

    # endregion Overrides

    # region math and comparison
    def __int__(self) -> int:
        return round(self.value)

    def __eq__(self, other: object) -> bool:
        if isinstance(other, UnitInch10):
            return self.almost_equal(other.value)
        if hasattr(other, "get_value_inch10"):
            oth_val = other.get_value_inch10()  # type: ignore
            return self.almost_equal(oth_val)
        if hasattr(other, "get_value_mm100"):
            return self.get_value_mm100() == other.get_value_mm100()  # type: ignore
        with contextlib.suppress(Exception):
            return self.almost_equal(float(other))  # type: ignore
        return False

    def __add__(self, other: object) -> UnitInch10:
        if isinstance(other, UnitInch10):
            return self.from_inch10(self.value + other.value)
        if hasattr(other, "get_value_inch10"):
            oth_val = other.get_value_inch10()  # type: ignore
            return self.from_inch10(self.value + oth_val)
        if hasattr(other, "get_value_mm100"):
            oth_val = other.get_value_mm100()  # type: ignore
            oth_val_inch10 = UnitConvert.convert(num=oth_val, frm=UnitLength.MM100, to=UnitLength.IN10)
            return self.from_inch10(self.value + oth_val_inch10)  # type: ignore

        if isinstance(other, (int, float)):
            return self.from_inch10(self.value + other)  # type: ignore
        return NotImplemented

    def __radd__(self, other: object) -> UnitInch10:
        return self if other == 0 else self.__add__(other)

    def __sub__(self, other: object) -> UnitInch10:
        if isinstance(other, UnitInch10):
            return self.from_inch10(self.value - other.value)
        if hasattr(other, "get_value_inch10"):
            oth_val = other.get_value_inch10()  # type: ignore
            return self.from_inch10(self.value - oth_val)
        if hasattr(other, "get_value_mm100"):
            oth_val = other.get_value_mm100()  # type: ignore
            oth_val_inch10 = UnitConvert.convert(num=oth_val, frm=UnitLength.MM100, to=UnitLength.IN10)
            return self.from_inch10(self.value - oth_val_inch10)  # type: ignore

        if isinstance(other, (int, float)):
            return self.from_inch10(self.value - other)  # type: ignore
        return NotImplemented

    def __rsub__(self, other: object) -> UnitInch10:
        if isinstance(other, (int, float)):
            return self.from_inch10(other - self.value)  # type: ignore
        return NotImplemented

    def __mul__(self, other: object) -> UnitInch10:
        if isinstance(other, UnitInch10):
            return self.from_inch10(self.value * other.value)
        if hasattr(other, "get_value_inch10"):
            oth_val = other.get_value_inch10()  # type: ignore
            return self.from_inch10(self.value * oth_val)
        if hasattr(other, "get_value_mm100"):
            oth_val = other.get_value_mm100()  # type: ignore
            oth_val_inch10 = UnitConvert.convert(num=oth_val, frm=UnitLength.MM100, to=UnitLength.IN10)
            return self.from_inch10(self.value * oth_val_inch10)  # type: ignore

        if isinstance(other, (int, float)):
            return self.from_inch10(self.value * other)  # type: ignore

        return NotImplemented

    def __rmul__(self, other: int) -> UnitInch10:
        return self if other == 0 else self.__mul__(other)

    def __truediv__(self, other: object) -> UnitInch10:
        if isinstance(other, UnitInch10):
            if other.value == 0:
                raise ZeroDivisionError
            return self.from_inch10(self.value / other.value)
        if hasattr(other, "get_value_inch10"):
            oth_val = other.get_value_inch10()  # type: ignore
            if oth_val == 0:
                raise ZeroDivisionError
            return self.from_inch10(self.value / oth_val)
        if hasattr(other, "get_value_mm100"):
            oth_val = other.get_value_mm100()  # type: ignore
            oth_val_inch10 = UnitConvert.convert(num=oth_val, frm=UnitLength.MM100, to=UnitLength.IN10)
            if oth_val_inch10 == 0:
                raise ZeroDivisionError
            return self.from_inch10(self.value / oth_val_inch10)  # type: ignore
        if isinstance(other, (int, float)):
            if other == 0:
                raise ZeroDivisionError
            return self.from_inch10(self.value / other)  # type: ignore
        return NotImplemented

    def __rtruediv__(self, other: object) -> UnitInch10:
        if isinstance(other, (int, float)):
            if self.value == 0:
                raise ZeroDivisionError
            return self.from_inch10(other / self.value)  # type: ignore
        return NotImplemented

    def __abs__(self) -> float:
        return abs(self.value)

    def __lt__(self, other: object) -> bool:
        if isinstance(other, UnitInch10):
            return self.value < other.value
        if hasattr(other, "get_value_inch10"):
            oth_val = other.get_value_inch10()  # type: ignore
            return self.value < oth_val
        if hasattr(other, "get_value_mm100"):
            return self.get_value_mm100() < other.get_value_mm100()  # type: ignore
        with contextlib.suppress(Exception):
            return self.value < float(other)  # type: ignore
        return False

    def __le__(self, other: object) -> bool:
        if isinstance(other, UnitInch10):
            return True if self.almost_equal(other.value) else self.value < other.value
        if hasattr(other, "get_value_inch10"):
            oth_val = other.get_value_inch10()  # type: ignore
            return True if self.almost_equal(oth_val) else self.value < oth_val
        if hasattr(other, "get_value_mm100"):
            return self.get_value_mm100() <= other.get_value_mm100()  # type: ignore
        with contextlib.suppress(Exception):
            oth_val = float(other)  # type: ignore
            return True if self.almost_equal(oth_val) else self.value < oth_val
        return False

    def __gt__(self, other: object) -> bool:
        if isinstance(other, UnitInch10):
            return self.value > other.value
        if hasattr(other, "get_value_inch10"):
            oth_val = other.get_value_inch10()  # type: ignore
            return self.value > oth_val
        if hasattr(other, "get_value_mm100"):
            return self.get_value_mm100() > other.get_value_mm100()  # type: ignore
        with contextlib.suppress(Exception):
            return self.value > float(other)  # type: ignore
        return False

    def __ge__(self, other: object) -> bool:
        if isinstance(other, UnitInch10):
            return True if self.almost_equal(other.value) else self.value > other.value
        if hasattr(other, "get_value_inch10"):
            oth_val = other.get_value_inch10()  # type: ignore
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
            UnitLength: Instance unit length ``UnitLength.IN10``.
        """
        return UnitLength.IN10

    def convert_to(self, unit: UnitLength) -> float:
        """
        Converts instance value to specified unit.

        Args:
            unit (UnitLength): Unit to convert to.

        Returns:
            float: Value in specified unit.
        """
        return UnitConvert.convert(num=self.value, frm=UnitLength.IN10, to=unit)

    def get_value_cm(self) -> float:
        """
        Gets instance value converted to ``cm`` units.

        Returns:
            int: Value in ``cm`` units.
        """
        return UnitConvert.convert(num=self.value, frm=UnitLength.IN10, to=UnitLength.CM)

    def get_value_mm(self) -> float:
        """
        Gets instance value converted to ``mm`` units.

        Returns:
            int: Value in ``mm`` units.
        """
        return UnitConvert.convert(num=self.value, frm=UnitLength.IN10, to=UnitLength.MM)

    def get_value_mm100(self) -> int:
        """
        Gets instance value converted to ``1/100th mm`` units.

        Returns:
            int: Value in ``1/100th mm`` units.
        """
        return round(UnitConvert.convert(num=self.value, frm=UnitLength.IN10, to=UnitLength.MM100))

    def get_value_inch(self) -> float:
        """
        Gets instance value in inch units.

        Returns:
            float: Value in inch units.
        """
        return UnitConvert.convert(num=self.value, frm=UnitLength.IN10, to=UnitLength.IN)

    def get_value_inch10(self) -> float:
        """
        Gets instance value in ``1/10th inch`` units.

        Returns:
            float: Value in ``1/10th inch`` units.
        """
        return self.value

    def get_value_inch100(self) -> float:
        """
        Gets instance value in ``1/100th inch`` units.

        Returns:
            float: Value in ``1/100th inch`` units.
        """
        return UnitConvert.convert(num=self.value, frm=UnitLength.IN10, to=UnitLength.IN100)

    def get_value_inch1000(self) -> int:
        """
        Gets instance value in ``1/1000th inch`` units.

        Returns:
            int: Value in ``1/100th inch`` units.
        """
        return round(UnitConvert.convert(num=self.value, frm=UnitLength.IN10, to=UnitLength.IN1000))

    def get_value_pt(self) -> float:
        """
        Gets instance value converted to Size in ``pt`` (points) units.

        Returns:
            int: Value in ``pt`` units.
        """
        return UnitConvert.convert(num=self.value, frm=UnitLength.IN10, to=UnitLength.PT)

    def get_value_px(self) -> float:
        """
        Gets instance value in ``px`` (pixel) units.

        Returns:
            int: Value in ``px`` units.
        """
        return UnitConvert.convert(num=self.value, frm=UnitLength.IN10, to=UnitLength.PX)

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
        af = af_type.from_inch10(self.value)
        return af.value

    @classmethod
    def from_mm100(cls: Type[_TUnitInch10], value: int) -> _TUnitInch10:
        """
        Get instance from ``1/100th mm`` value.

        Args:
            value (int): ``1/100th mm`` value.

        Returns:
            UnitInch10:
        """
        # pylint: disable=unnecessary-dunder-call
        inst = super(UnitInch10, cls).__new__(cls)  # type: ignore
        inst.__init__(UnitConvert.convert(num=value, frm=UnitLength.MM100, to=UnitLength.IN10))
        return inst

    @classmethod
    def from_pt(cls: Type[_TUnitInch10], value: float) -> _TUnitInch10:
        """
        Get instance from ``pt`` (points) value.

        Args:
            value (float): ``pt`` value.

        Returns:
            UnitInch10:
        """
        # pylint: disable=unnecessary-dunder-call
        inst = super(UnitInch10, cls).__new__(cls)  # type: ignore
        inst.__init__(float(UnitConvert.convert(num=value, frm=UnitLength.PT, to=UnitLength.IN10)))
        return inst

    @classmethod
    def from_px(cls: Type[_TUnitInch10], value: float) -> _TUnitInch10:
        """
        Get instance from ``px`` (pixel) value.

        Args:
            value (float): ``px`` value.

        Returns:
            UnitInch10:
        """
        # pylint: disable=unnecessary-dunder-call
        inst = super(UnitInch10, cls).__new__(cls)  # type: ignore
        inst.__init__(UnitConvert.convert(num=value, frm=UnitLength.PX, to=UnitLength.IN10))
        return inst

    @classmethod
    def from_inch(cls: Type[_TUnitInch10], value: float) -> _TUnitInch10:
        """
        Get instance from ``in`` (inch) value.

        Args:
            value (float): ``in`` value.

        Returns:
            UnitInch10:
        """
        # pylint: disable=unnecessary-dunder-call
        inst = super(UnitInch10, cls).__new__(cls)  # type: ignore
        inst.__init__(float(UnitConvert.convert(num=value, frm=UnitLength.IN, to=UnitLength.IN10)))
        return inst

    from_in = from_inch

    @classmethod
    def from_inch10(cls: Type[_TUnitInch10], value: float) -> _TUnitInch10:
        """
        Get instance from ``1/10th in`` (inch) value.

        Args:
            value (int): ``1/10th in`` value.

        Returns:
            UnitInch10:
        """
        # pylint: disable=unnecessary-dunder-call
        inst = super(UnitInch10, cls).__new__(cls)  # type: ignore
        inst.__init__(value)
        return inst

    @classmethod
    def from_inch100(cls: Type[_TUnitInch10], value: float) -> _TUnitInch10:
        """
        Get instance from ``1/100th in`` (inch) value.

        Args:
            value (int): ``1/100th in`` value.

        Returns:
            UnitInch10:
        """
        # pylint: disable=unnecessary-dunder-call
        inst = super(UnitInch10, cls).__new__(cls)  # type: ignore
        inst.__init__(UnitConvert.convert(num=value, frm=UnitLength.IN100, to=UnitLength.IN10))
        return inst

    @classmethod
    def from_inch1000(cls: Type[_TUnitInch10], value: float) -> _TUnitInch10:
        """
        Get instance from ``1/1,000th in`` (inch) value.

        Args:
            value (int): ``1/1,000th in`` value.

        Returns:
            UnitInch10:
        """
        # pylint: disable=unnecessary-dunder-call
        inst = super(UnitInch10, cls).__new__(cls)  # type: ignore
        inst.__init__(UnitConvert.convert(num=value, frm=UnitLength.IN1000, to=UnitLength.IN10))
        return inst

    @classmethod
    def from_cm(cls: Type[_TUnitInch10], value: float) -> _TUnitInch10:
        """
        Get instance from ``cm`` value.

        Args:
            value (float): ``cm`` value.

        Returns:
            UnitInch10:
        """
        # pylint: disable=unnecessary-dunder-call
        inst = super(UnitInch10, cls).__new__(cls)  # type: ignore
        inst.__init__(UnitConvert.convert(num=value, frm=UnitLength.CM, to=UnitLength.IN10))
        return inst

    @classmethod
    def from_app_font(cls: Type[_TUnitInch10], value: float, kind: PointSizeKind | int) -> _TUnitInch10:
        """
        Get instance from ``AppFont`` value.

        Args:
            value (int): ``AppFont`` value.
            kind (PointSizeKind, optional): The kind of ``AppFont`` to use.

        Returns:
            UnitInch10:

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
        return cls.from_inch10(af.get_value_inch10())

    @classmethod
    def from_unit_val(cls: Type[_TUnitInch10], value: UnitT | float | int) -> _TUnitInch10:
        """
        Get instance from ``UnitT`` or float value.

        Args:
            value (UnitT, float, int): ``UnitT`` or float value. If float then it is assumed to be in ``inch10`` units.

        Returns:
            UnitInch10:
        """
        try:
            if hasattr(value, "get_value_inch10"):
                return cls.from_inch10(value.get_value_inch10())  # type: ignore

            unit_100 = value.get_value_mm100()  # type: ignore
            return cls.from_mm100(unit_100)
        except AttributeError:
            return cls.from_inch10(float(value))  # type: ignore
