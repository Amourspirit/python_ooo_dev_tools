from __future__ import annotations
import contextlib
from typing import TypeVar, Type, TYPE_CHECKING
from dataclasses import dataclass

from ooodev.utils.decorator import enforce
from ooodev.units.unit_convert import UnitConvert
from ooodev.units.unit_convert import UnitLength
from ooodev.utils.kind.point_size_kind import PointSizeKind

if TYPE_CHECKING:
    from ooodev.units.unit_obj import UnitT

_TUnitMM100 = TypeVar("_TUnitMM100", bound="UnitMM100")


# Note that from __future__ import annotations converts annotations to string.
# this means that @enforce.enforce_types will see string as type. This is fine in
# most cases. Especially for built in types.
@enforce.enforce_types
@dataclass(unsafe_hash=True)
class UnitMM100:
    """
    Represents ``1/100th mm`` units.

    Supports ``UnitT`` protocol.

    See Also:
        :ref:`proto_unit_obj`
    """

    value: int
    """Int value."""

    # region math and comparison
    def __int__(self) -> int:
        return self.value

    def __eq__(self, other: object) -> bool:
        if hasattr(other, "get_value_mm100"):
            return self.get_value_mm100() == other.get_value_mm100()  # type: ignore
        with contextlib.suppress(Exception):
            return self.get_value_mm100() == int(other)  # type: ignore
        return False

    def __add__(self, other: object) -> UnitMM100:
        if hasattr(other, "get_value_mm100"):
            oth_val = other.get_value_mm100()  # type: ignore
            return self.from_mm100(self.value + oth_val)  # type: ignore

        if isinstance(other, (int, float)):
            return self.from_mm100(self.value + int(other))  # type: ignore

        return NotImplemented

    def __radd__(self, other: object) -> UnitMM100:
        return self if other == 0 else self.__add__(other)

    def __sub__(self, other: object) -> UnitMM100:
        if hasattr(other, "get_value_mm100"):
            oth_val = other.get_value_mm100()  # type: ignore
            return self.from_mm100(self.value - oth_val)  # type: ignore

        if isinstance(other, (int, float)):
            return self.from_mm100(self.value - int(other))  # type: ignore
        return NotImplemented

    def __rsub__(self, other: object) -> UnitMM100:
        if isinstance(other, (int, float)):
            self_val = self.get_value_mm100()
            return self.from_mm100(int(other) - self_val)  # type: ignore
        return NotImplemented

    def __mul__(self, other: object) -> UnitMM100:
        if hasattr(other, "get_value_mm100"):
            oth_val = other.get_value_mm100()  # type: ignore
            return self.from_mm100(self.value * oth_val)  # type: ignore

        if isinstance(other, (int, float)):
            return self.from_mm100(self.value * int(other))  # type: ignore

        return NotImplemented

    def __rmul__(self, other: int) -> UnitMM100:
        return self if other == 0 else self.__mul__(other)

    def __truediv__(self, other: object) -> UnitMM100:
        if hasattr(other, "get_value_mm100"):
            oth_val = other.get_value_mm100()  # type: ignore
            if oth_val == 0:
                raise ZeroDivisionError
            return self.from_mm100(self.value // oth_val)  # type: ignore
        if isinstance(other, (int, float)):
            if other == 0:
                raise ZeroDivisionError
            return self.from_mm100(self.value // other)  # type: ignore
        return NotImplemented

    def __rtruediv__(self, other: object) -> UnitMM100:
        if isinstance(other, (int, float)):
            if self.value == 0:
                raise ZeroDivisionError
            return self.from_mm100(other // self.value)  # type: ignore
        return NotImplemented

    def __abs__(self) -> int:
        return abs(self.value)

    def __lt__(self, other: object) -> bool:
        if hasattr(other, "get_value_mm100"):
            return self.value < other.get_value_mm100()  # type: ignore
        with contextlib.suppress(Exception):
            return self.value < int(other)  # type: ignore
        return False

    def __le__(self, other: object) -> bool:
        if hasattr(other, "get_value_mm100"):
            return self.value <= other.get_value_mm100()  # type: ignore
        with contextlib.suppress(Exception):
            return self.value <= int(other)  # type: ignore
        return False

    def __gt__(self, other: object) -> bool:
        if hasattr(other, "get_value_mm100"):
            return self.value > other.get_value_mm100()  # type: ignore
        with contextlib.suppress(Exception):
            return self.value > int(other)  # type: ignore
        return False

    def __ge__(self, other: object) -> bool:
        if hasattr(other, "get_value_mm100"):
            return self.value >= other.get_value_mm100()  # type: ignore
        with contextlib.suppress(Exception):
            return self.value >= int(other)  # type: ignore
        return False

    # endregion math and comparison

    @staticmethod
    def get_unit_length() -> UnitLength:
        """
        Gets instance unit length.

        Returns:
            UnitLength: Instance unit length ``UnitLength.MM100``.
        """
        return UnitLength.MM100

    def convert_to(self, unit: UnitLength) -> float:
        """
        Converts instance value to specified unit.

        Args:
            unit (UnitLength): Unit to convert to.

        Returns:
            float: Value in specified unit.
        """
        return UnitConvert.convert(num=self.value, frm=UnitLength.MM100, to=unit)

    def get_value_cm(self) -> float:
        """
        Gets instance value converted to ``cm`` units.

        Returns:
            int: Value in ``cm`` units.
        """
        return UnitConvert.convert(num=self.value, frm=UnitLength.MM100, to=UnitLength.CM)

    def get_value_mm(self) -> float:
        """
        Gets instance value converted to ``mm`` units.

        Returns:
            int: Value in ``mm`` units.
        """
        return UnitConvert.convert_mm100_mm(self.value)

    def get_value_mm10(self) -> float:
        """
        Gets instance value in ``1/10th mm`` units.

        Returns:
            float: Value in ``1/10th mm`` units.
        """
        return UnitConvert.convert(num=self.value, frm=UnitLength.MM100, to=UnitLength.MM10)

    def get_value_mm100(self) -> int:
        """
        Gets instance value in ``1/100th mm`` units.

        Returns:
            int: Value in ``1/100th mm`` units.
        """
        return self.value

    def get_value_pt(self) -> float:
        """
        Gets instance value converted to ``pt`` (point) units.

        Returns:
            int: Value in ``pt`` units.
        """
        return UnitConvert.convert(num=self.value, frm=UnitLength.MM100, to=UnitLength.PT)

    def get_value_px(self) -> float:
        """
        Gets instance value in ``px`` (pixel) units.

        Returns:
            int: Value in ``px`` units.
        """
        return UnitConvert.convert(num=self.value, frm=UnitLength.MM100, to=UnitLength.PX)

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
        af = af_type.from_mm100(self.value)
        return af.value

    @classmethod
    def from_mm(cls: Type[_TUnitMM100], value: float) -> _TUnitMM100:
        """
        Get instance from ``mm`` value.

        Args:
            value (int): ``mm`` value.

        Returns:
            UnitMM100:
        """
        # pylint: disable=unnecessary-dunder-call
        inst = super(UnitMM100, cls).__new__(cls)  # type: ignore
        inst.__init__(UnitConvert.convert_mm_mm100(value))
        return inst

    @classmethod
    def from_mm10(cls: Type[_TUnitMM100], value: float) -> _TUnitMM100:
        """
        Get instance from ``1/10th mm`` value.

        Args:
            value (int): ``1/10th mm`` value.

        Returns:
            UnitMM100:
        """
        # pylint: disable=unnecessary-dunder-call
        inst = super(UnitMM100, cls).__new__(cls)  # type: ignore
        inst.__init__(round(UnitConvert.convert(num=value, frm=UnitLength.MM10, to=UnitLength.MM100)))
        return inst

    @classmethod
    def from_mm100(cls: Type[_TUnitMM100], value: int) -> _TUnitMM100:
        """
        Get instance from ``1/100th mm`` value.

        Args:
            value (int): ``1/100th mm`` value.

        Returns:
            UnitMM100:
        """
        # sourcery skip: remove-unnecessary-cast
        # pylint: disable=unnecessary-dunder-call
        inst = super(UnitMM100, cls).__new__(cls)  # type: ignore
        inst.__init__(int(value))
        return inst

    @classmethod
    def from_pt(cls: Type[_TUnitMM100], value: float) -> _TUnitMM100:
        """
        Get instance from ``pt`` (points) value.

        Args:
            value (int): ``pt`` value.

        Returns:
            UnitMM100:
        """
        # pylint: disable=unnecessary-dunder-call
        inst = super(UnitMM100, cls).__new__(cls)  # type: ignore
        inst.__init__(round(UnitConvert.convert(num=value, frm=UnitLength.PT, to=UnitLength.MM100)))
        return inst

    @classmethod
    def from_px(cls: Type[_TUnitMM100], value: float) -> _TUnitMM100:
        """
        Get instance from ``px`` (pixel) value.

        Args:
            value (float): ``px`` value.

        Returns:
            UnitMM100:
        """
        # pylint: disable=unnecessary-dunder-call
        inst = super(UnitMM100, cls).__new__(cls)  # type: ignore
        inst.__init__(round(UnitConvert.convert(num=value, frm=UnitLength.PX, to=UnitLength.MM100)))
        return inst

    @classmethod
    def from_inch(cls: Type[_TUnitMM100], value: float) -> _TUnitMM100:
        """
        Get instance from ``in`` (inch) value.

        Args:
            value (int): ``in`` value.

        Returns:
            UnitMM100:
        """
        # pylint: disable=unnecessary-dunder-call
        inst = super(UnitMM100, cls).__new__(cls)  # type: ignore
        inst.__init__(round(UnitConvert.convert(num=value, frm=UnitLength.IN, to=UnitLength.MM100)))
        return inst

    @classmethod
    def from_inch10(cls: Type[_TUnitMM100], value: float) -> _TUnitMM100:
        """
        Get instance from ``1/10th in`` (inch) value.

        Args:
            value (int): ```/10th in`` value.

        Returns:
            UnitMM100:
        """
        # pylint: disable=unnecessary-dunder-call
        inst = super(UnitMM100, cls).__new__(cls)  # type: ignore
        inst.__init__(round(UnitConvert.convert(num=value, frm=UnitLength.IN10, to=UnitLength.MM100)))
        return inst

    @classmethod
    def from_inch100(cls: Type[_TUnitMM100], value: float) -> _TUnitMM100:
        """
        Get instance from ``1/100th in`` (inch) value.

        Args:
            value (int): ``1/100th in`` value.

        Returns:
            UnitMM100:
        """
        # pylint: disable=unnecessary-dunder-call
        inst = super(UnitMM100, cls).__new__(cls)  # type: ignore
        inst.__init__(round(UnitConvert.convert(num=value, frm=UnitLength.IN100, to=UnitLength.MM100)))
        return inst

    @classmethod
    def from_inch1000(cls: Type[_TUnitMM100], value: int) -> _TUnitMM100:
        """
        Get instance from ``1/1,000th in`` (inch) value.

        Args:
            value (int): ``1/1,000th in`` value.

        Returns:
            UnitMM100:
        """
        # pylint: disable=unnecessary-dunder-call
        inst = super(UnitMM100, cls).__new__(cls)  # type: ignore
        inst.__init__(round(UnitConvert.convert(num=value, frm=UnitLength.IN1000, to=UnitLength.MM100)))
        return inst

    @classmethod
    def from_cm(cls: Type[_TUnitMM100], value: float) -> _TUnitMM100:
        """
        Get instance from ``cm`` value.

        Args:
            value (int): ``cm`` value.

        Returns:
            UnitMM100:
        """
        # pylint: disable=unnecessary-dunder-call
        inst = super(UnitMM100, cls).__new__(cls)  # type: ignore
        inst.__init__(round(UnitConvert.convert(num=value, frm=UnitLength.CM, to=UnitLength.MM100)))
        return inst

    @classmethod
    def from_app_font(cls: Type[_TUnitMM100], value: float, kind: PointSizeKind | int) -> _TUnitMM100:
        """
        Get instance from ``AppFont`` value.

        Args:
            value (int): ``AppFont`` value.
            kind (PointSizeKind, optional): The kind of ``AppFont`` to use.

        Returns:
            UnitMM100:

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
        return cls.from_mm100(af.get_value_mm100())

    @classmethod
    def from_unit_val(cls: Type[_TUnitMM100], value: UnitT | float | int) -> _TUnitMM100:
        """
        Get instance from ``UnitT`` or float value.

        Args:
            value (UnitT, float, int): ``UnitT`` or float value. If float then it is assumed to be in ``1/100th mm`` units.

        Returns:
            UnitMM100:
        """
        try:
            unit_100 = value.get_value_mm100()  # type: ignore
            return cls.from_mm100(unit_100)
        except AttributeError:
            return cls.from_mm100(round(value))  # type: ignore
