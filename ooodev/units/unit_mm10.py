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

_TUnitMM10 = TypeVar("_TUnitMM10", bound="UnitMM10")


@dataclass(unsafe_hash=True)
class UnitMM10(BaseFloatValue):
    """
    Unit in ``1/10th mm`` units.

    Supports ``UnitT`` protocol.

    See Also:
        :ref:`proto_unit_obj`
    """

    def __post_init__(self):
        if not isinstance(self.value, float):
            object.__setattr__(self, "value", float(self.value))

    # region Overrides
    def _from_float(self, value: float) -> UnitMM10:
        return self.from_mm10(value)

    # endregion Overrides

    # region math and comparison
    @override
    def __int__(self) -> int:
        return round(self.value)

    def __eq__(self, other: object) -> bool:
        if isinstance(other, UnitMM10):
            return self.almost_equal(other.value)
        if hasattr(other, "get_value_mm10"):
            oth_val = other.get_value_mm10()  # type: ignore
            return self.almost_equal(oth_val)
        if hasattr(other, "get_value_mm100"):
            return self.get_value_mm100() == other.get_value_mm100()  # type: ignore
        with contextlib.suppress(Exception):
            return self.almost_equal(float(other))  # type: ignore
        return False

    @override
    def __add__(self, other: object) -> UnitMM10:
        if isinstance(other, UnitMM10):
            return self.from_mm10(self.value + other.value)
        if hasattr(other, "get_value_mm10"):
            oth_val = other.get_value_mm10()  # type: ignore
            return self.from_mm10(self.value + oth_val)
        if hasattr(other, "get_value_mm100"):
            oth_val = other.get_value_mm100()  # type: ignore
            oth_val_mm10 = UnitConvert.convert(num=oth_val, frm=UnitLength.MM100, to=UnitLength.MM10)
            return self.from_mm10(self.value + oth_val_mm10)  # type: ignore

        if isinstance(other, (int, float)):
            return self.from_mm10(self.value + other)  # type: ignore
        return NotImplemented

    @override
    def __radd__(self, other: object) -> UnitMM10:
        return self if other == 0 else self.__add__(other)

    @override
    def __sub__(self, other: object) -> UnitMM10:
        if isinstance(other, UnitMM10):
            return self.from_mm10(self.value - other.value)
        if hasattr(other, "get_value_mm10"):
            oth_val = other.get_value_mm10()  # type: ignore
            return self.from_mm10(self.value - oth_val)
        if hasattr(other, "get_value_mm100"):
            oth_val = other.get_value_mm100()  # type: ignore
            oth_val_mm10 = UnitConvert.convert(num=oth_val, frm=UnitLength.MM100, to=UnitLength.MM10)
            return self.from_mm10(self.value - oth_val_mm10)  # type: ignore

        if isinstance(other, (int, float)):
            return self.from_mm10(self.value - other)  # type: ignore
        return NotImplemented

    @override
    def __rsub__(self, other: object) -> UnitMM10:
        if isinstance(other, (int, float)):
            return self.from_mm10(other - self.value)  # type: ignore
        return NotImplemented

    @override
    def __mul__(self, other: object) -> UnitMM10:
        if isinstance(other, UnitMM10):
            return self.from_mm10(self.value * other.value)
        if hasattr(other, "get_value_mm10"):
            oth_val = other.get_value_mm10()  # type: ignore
            return self.from_mm10(self.value * oth_val)
        if hasattr(other, "get_value_mm100"):
            oth_val = other.get_value_mm100()  # type: ignore
            oth_val_mm10 = UnitConvert.convert(num=oth_val, frm=UnitLength.MM100, to=UnitLength.MM10)
            return self.from_mm10(self.value * oth_val_mm10)  # type: ignore

        if isinstance(other, (int, float)):
            return self.from_mm10(self.value * other)  # type: ignore

        return NotImplemented

    @override
    def __rmul__(self, other: int) -> UnitMM10:
        return self if other == 0 else self.__mul__(other)

    @override
    def __truediv__(self, other: object) -> UnitMM10:
        if isinstance(other, UnitMM10):
            if other.value == 0:
                raise ZeroDivisionError
            return self.from_mm10(self.value / other.value)
        if hasattr(other, "get_value_mm10"):
            oth_val = other.get_value_mm10()  # type: ignore
            if oth_val == 0:
                raise ZeroDivisionError
            return self.from_mm10(self.value / oth_val)
        if hasattr(other, "get_value_mm100"):
            oth_val = other.get_value_mm100()  # type: ignore
            oth_val_mm10 = UnitConvert.convert(num=oth_val, frm=UnitLength.MM100, to=UnitLength.MM10)
            if oth_val_mm10 == 0:
                raise ZeroDivisionError
            return self.from_mm10(self.value / oth_val_mm10)  # type: ignore
        if isinstance(other, (int, float)):
            if other == 0:
                raise ZeroDivisionError
            return self.from_mm10(self.value / other)  # type: ignore
        return NotImplemented

    @override
    def __rtruediv__(self, other: object) -> UnitMM10:
        if isinstance(other, (int, float)):
            if self.value == 0:
                raise ZeroDivisionError
            return self.from_mm10(other / self.value)  # type: ignore
        return NotImplemented

    @override
    def __abs__(self) -> float:  # type: ignore
        return abs(self.value)

    @override
    def __lt__(self, other: object) -> bool:
        if isinstance(other, UnitMM10):
            return self.value < other.value
        if hasattr(other, "get_value_mm10"):
            oth_val = other.get_value_mm10()  # type: ignore
            return self.value < oth_val
        if hasattr(other, "get_value_mm100"):
            return self.get_value_mm100() < other.get_value_mm100()  # type: ignore
        with contextlib.suppress(Exception):
            return self.value < float(other)  # type: ignore
        return False

    @override
    def __le__(self, other: object) -> bool:
        if isinstance(other, UnitMM10):
            return True if self.almost_equal(other.value) else self.value < other.value
        if hasattr(other, "get_value_mm10"):
            oth_val = other.get_value_mm10()  # type: ignore
            return True if self.almost_equal(oth_val) else self.value < oth_val
        if hasattr(other, "get_value_mm100"):
            return self.get_value_mm100() <= other.get_value_mm100()  # type: ignore
        with contextlib.suppress(Exception):
            oth_val = float(other)  # type: ignore
            return True if self.almost_equal(oth_val) else self.value < oth_val
        return False

    @override
    def __gt__(self, other: object) -> bool:
        if isinstance(other, UnitMM10):
            return self.value > other.value
        if hasattr(other, "get_value_mm10"):
            oth_val = other.get_value_mm10()  # type: ignore
            return self.value > oth_val
        if hasattr(other, "get_value_mm100"):
            return self.get_value_mm100() > other.get_value_mm100()  # type: ignore
        with contextlib.suppress(Exception):
            return self.value > float(other)  # type: ignore
        return False

    @override
    def __ge__(self, other: object) -> bool:
        if isinstance(other, UnitMM10):
            return True if self.almost_equal(other.value) else self.value > other.value
        if hasattr(other, "get_value_mm10"):
            oth_val = other.get_value_mm10()  # type: ignore
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
            UnitLength: Instance unit length ``UnitLength.MM10``.
        """
        return UnitLength.MM10

    def convert_to(self, unit: UnitLength) -> float:
        """
        Converts instance value to specified unit.

        Args:
            unit (UnitLength): Unit to convert to.

        Returns:
            float: Value in specified unit.
        """
        return UnitConvert.convert(num=self.value, frm=UnitLength.MM10, to=unit)

    def get_value_cm(self) -> float:
        """
        Gets instance value converted to ``cm`` units.

        Returns:
            int: Value in ``cm`` units.
        """
        return UnitConvert.convert(num=self.value, frm=UnitLength.MM10, to=UnitLength.CM)

    def get_value_mm(self) -> float:
        """
        Gets instance value converted to ``mm`` units.

        Returns:
            int: Value in ``mm`` units.
        """
        return UnitConvert.convert(num=self.value, frm=UnitLength.MM10, to=UnitLength.MM)

    def get_value_mm10(self) -> float:
        """
        Gets instance value in ``1/10th mm`` units.

        Returns:
            float: Value in ``1/10th mm`` units.
        """
        return self.value

    def get_value_mm100(self) -> int:
        """
        Gets instance value converted to ``1/100th mm`` units.

        Returns:
            int: Value in ``1/100th mm`` units.
        """
        return round(UnitConvert.convert(num=self.value, frm=UnitLength.MM10, to=UnitLength.MM100))

    def get_value_pt(self) -> float:
        """
        Gets instance value converted to ``pt`` (points) units.

        Returns:
            int: Value in ``pt`` units.
        """
        return UnitConvert.convert(num=self.value, frm=UnitLength.MM10, to=UnitLength.PT)

    def get_value_px(self) -> float:
        """
        Gets instance value in ``px`` (pixel) units.

        Returns:
            int: Value in ``px`` units.
        """
        return UnitConvert.convert(num=self.value, frm=UnitLength.MM10, to=UnitLength.PX)

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
        af = af_type.from_mm10(self.value)
        return af.value

    @classmethod
    def from_pt(cls: Type[_TUnitMM10], value: float) -> _TUnitMM10:
        """
        Get instance from ``pt`` (points) value.

        Args:
            value (float): ``pt`` value.

        Returns:
            UnitMM10:
        """
        # pylint: disable=unnecessary-dunder-call
        inst = super(UnitMM10, cls).__new__(cls)  # type: ignore
        inst.__init__(float(UnitConvert.convert(num=value, frm=UnitLength.PT, to=UnitLength.MM10)))
        return inst

    @classmethod
    def from_px(cls: Type[_TUnitMM10], value: float) -> _TUnitMM10:
        """
        Get instance from ``px`` (pixel) value.

        Args:
            value (float): ``px`` value.

        Returns:
            UnitMM10:
        """
        # pylint: disable=unnecessary-dunder-call
        inst = super(UnitMM10, cls).__new__(cls)  # type: ignore
        inst.__init__(UnitConvert.convert(num=value, frm=UnitLength.PX, to=UnitLength.MM10))
        return inst

    @classmethod
    def from_mm(cls: Type[_TUnitMM10], value: float) -> _TUnitMM10:
        """
        Get instance from ``mm`` value.

        Args:
            value (int): ``mm`` value.

        Returns:
            UnitMM10:
        """
        # pylint: disable=unnecessary-dunder-call
        inst = super(UnitMM10, cls).__new__(cls)  # type: ignore
        inst.__init__(UnitConvert.convert(num=value, frm=UnitLength.MM, to=UnitLength.MM10))
        return inst

    @classmethod
    def from_mm10(cls: Type[_TUnitMM10], value: float) -> _TUnitMM10:
        """
        Get instance from ``1/10th mm`` value.

        Args:
            value (int): ``1/10th mm`` value.

        Returns:
            UnitMM10:
        """
        # pylint: disable=unnecessary-dunder-call
        inst = super(UnitMM10, cls).__new__(cls)  # type: ignore
        inst.__init__(value)
        return inst

    @classmethod
    def from_mm100(cls: Type[_TUnitMM10], value: int) -> _TUnitMM10:
        """
        Get instance from ``1/100th mm`` value.

        Args:
            value (int): ``1/100th mm`` value.

        Returns:
            UnitMM10:
        """
        # pylint: disable=unnecessary-dunder-call
        inst = super(UnitMM10, cls).__new__(cls)  # type: ignore
        inst.__init__(UnitConvert.convert(num=value, frm=UnitLength.MM100, to=UnitLength.MM10))
        return inst

    @classmethod
    def from_inch(cls: Type[_TUnitMM10], value: float) -> _TUnitMM10:
        """
        Get instance from ``in`` (inch) value.

        Args:
            value (int): ``in`` value.

        Returns:
            UnitMM10:
        """
        # pylint: disable=unnecessary-dunder-call
        inst = super(UnitMM10, cls).__new__(cls)  # type: ignore
        inst.__init__(UnitConvert.convert(num=value, frm=UnitLength.IN, to=UnitLength.MM10))
        return inst

    @classmethod
    def from_inch10(cls: Type[_TUnitMM10], value: float) -> _TUnitMM10:
        """
        Get instance from ``1/10th in`` (inch) value.

        Args:
            value (int): ```/10th in`` value.

        Returns:
            UnitMM10:
        """
        # pylint: disable=unnecessary-dunder-call
        inst = super(UnitMM10, cls).__new__(cls)  # type: ignore
        inst.__init__(UnitConvert.convert(num=value, frm=UnitLength.IN10, to=UnitLength.MM10))
        return inst

    @classmethod
    def from_inch100(cls: Type[_TUnitMM10], value: float) -> _TUnitMM10:
        """
        Get instance from ``1/100th in`` (inch) value.

        Args:
            value (int): ``1/100th in`` value.

        Returns:
            UnitMM10:
        """
        # pylint: disable=unnecessary-dunder-call
        inst = super(UnitMM10, cls).__new__(cls)  # type: ignore
        inst.__init__(UnitConvert.convert(num=value, frm=UnitLength.IN100, to=UnitLength.MM10))
        return inst

    @classmethod
    def from_inch1000(cls: Type[_TUnitMM10], value: int) -> _TUnitMM10:
        """
        Get instance from ``1/1,000th in`` (inch) value.

        Args:
            value (int): ``1/1,000th in`` value.

        Returns:
            UnitMM10:
        """
        # pylint: disable=unnecessary-dunder-call
        inst = super(UnitMM10, cls).__new__(cls)  # type: ignore
        inst.__init__(UnitConvert.convert(num=value, frm=UnitLength.IN1000, to=UnitLength.MM10))
        return inst

    @classmethod
    def from_cm(cls: Type[_TUnitMM10], value: float) -> _TUnitMM10:
        """
        Get instance from ``cm`` value.

        Args:
            value (int): ``cm`` value.

        Returns:
            UnitMM10:
        """
        # pylint: disable=unnecessary-dunder-call
        inst = super(UnitMM10, cls).__new__(cls)  # type: ignore
        inst.__init__(UnitConvert.convert(num=value, frm=UnitLength.CM, to=UnitLength.MM10))
        return inst

    @classmethod
    def from_app_font(cls: Type[_TUnitMM10], value: float, kind: PointSizeKind | int) -> _TUnitMM10:
        """
        Get instance from ``AppFont`` value.

        Args:
            value (int): ``AppFont`` value.
            kind (PointSizeKind, optional): The kind of ``AppFont`` to use.

        Returns:
            UnitMM10:

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
    def from_unit_val(cls: Type[_TUnitMM10], value: UnitT | float | int) -> _TUnitMM10:
        """
        Get instance from ``UnitT`` or float value.

        Args:
            value (UnitT, float, int): ``UnitT`` or float value. If float then it is assumed to be in ``mm`` units.

        Returns:
            UnitMM10:
        """
        try:
            if hasattr(value, "get_value_mm10"):
                return cls.from_mm10(value.get_value_mm10())  # type: ignore

            unit_val = value.get_value_mm()  # type: ignore
            return cls.from_mm(unit_val)
        except AttributeError:
            return cls.from_mm10(float(value))  # type: ignore
