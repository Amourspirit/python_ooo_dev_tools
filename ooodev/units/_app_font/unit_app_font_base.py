from __future__ import annotations
from abc import abstractmethod
import contextlib
from typing import Any, TYPE_CHECKING
from dataclasses import dataclass, field
from ooodev.utils.data_type.base_float_value import BaseFloatValue
from ooodev.units.unit_convert import UnitConvert
from ooodev.units.unit_convert import UnitLength
from ooodev.utils.kind.point_size_kind import PointSizeKind

if TYPE_CHECKING:
    from typing_extensions import Self
    from ooodev.units.unit_obj import UnitT
else:
    Self = Any


@dataclass(unsafe_hash=True)
class UnitAppFontBase(BaseFloatValue):
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

        if not isinstance(self.value, float):
            object.__setattr__(self, "value", float(self.value))
        self._set_ratio()

    @abstractmethod
    def _set_ratio(self) -> None:
        raise NotImplementedError

    # region Overrides
    def _from_float(self, value: float) -> Self:
        return self.from_app_font(value)

    @abstractmethod
    def get_app_font_kind(self) -> PointSizeKind:
        raise NotImplementedError

    # endregion Overrides

    # region math and comparison
    def __int__(self) -> int:
        return round(self.value)

    def __eq__(self, other: object) -> bool:
        if isinstance(other, UnitAppFontBase):
            # AppFontX and AppFontY are not equal
            if other.__class__.__name__ == self.__class__.__name__:
                return self.almost_equal(other.value)
            return self.get_value_mm100() == other.get_value_mm100()
        if hasattr(other, "get_value_app_font"):
            oth_val = other.get_value_app_font(kind=self.get_app_font_kind())  # type: ignore
            return self.almost_equal(oth_val)
        if hasattr(other, "get_value_mm100"):
            return self.get_value_mm100() == other.get_value_mm100()  # type: ignore
        with contextlib.suppress(Exception):
            return self.almost_equal(float(other))  # type: ignore
        return False

    def __add__(self, other: object) -> Self:
        if isinstance(other, UnitAppFontBase) and other.__class__.__name__ == self.__class__.__name__:
            return self.from_app_font(self.value + other.value)
        if hasattr(other, "get_value_px"):
            oth_val = other.get_value_px()  # type: ignore
            return self.from_px(self.get_value_px() + oth_val)
        if hasattr(other, "get_value_mm100"):
            oth_val = other.get_value_mm100()  # type: ignore
            oth_val_px = UnitConvert.convert(num=oth_val, frm=UnitLength.MM100, to=UnitLength.PX)
            return self.from_px(self.get_value_px() + oth_val_px)  # type: ignore

        if isinstance(other, (int, float)):
            return self.from_app_font(self.value + other)  # type: ignore
        return NotImplemented

    def __radd__(self, other: object) -> Self:
        return self if other == 0 else self.__add__(other)

    def __sub__(self, other: object) -> Self:
        if isinstance(other, UnitAppFontBase) and other.__class__.__name__ == self.__class__.__name__:
            return self.from_app_font(self.value - other.value)
        if hasattr(other, "get_value_px"):
            oth_val = other.get_value_px()  # type: ignore
            return self.from_px(self.get_value_px() - oth_val)
        if hasattr(other, "get_value_mm100"):
            oth_val = other.get_value_mm100()  # type: ignore
            oth_val_px = UnitConvert.convert(num=oth_val, frm=UnitLength.MM100, to=UnitLength.PX)
            return self.from_px(self.get_value_px() - oth_val_px)  # type: ignore

        if isinstance(other, (int, float)):
            return self.from_app_font(self.value - other)  # type: ignore
        return NotImplemented

    def __rsub__(self, other: object) -> Self:
        if isinstance(other, (int, float)):
            return self.from_app_font(other - self.value)  # type: ignore
        return NotImplemented

    def __mul__(self, other: object) -> Self:
        if isinstance(other, UnitAppFontBase) and other.__class__.__name__ == self.__class__.__name__:
            return self.from_app_font(self.value * other.value)
        if hasattr(other, "get_value_px"):
            oth_val = other.get_value_px()  # type: ignore
            # app_units = oth_val * self._ratio
            return self.from_px(self.get_value_px() * oth_val)
        if hasattr(other, "get_value_mm100"):
            oth_val = other.get_value_mm100()  # type: ignore
            oth_val_px = UnitConvert.convert(num=oth_val, frm=UnitLength.MM100, to=UnitLength.PX)
            return self.from_px(self.get_value_px() * oth_val_px)

        if isinstance(other, (int, float)):
            return self.from_app_font(self.value * other)  # type: ignore

        return NotImplemented

    def __rmul__(self, other: int) -> Self:
        return self if other == 0 else self.__mul__(other)

    def __truediv__(self, other: object) -> Self:
        if isinstance(other, UnitAppFontBase) and other.__class__.__name__ == self.__class__.__name__:
            if other.value == 0:
                raise ZeroDivisionError
            return self.from_app_font(self.value / other.value)
        if hasattr(other, "get_value_px"):
            oth_val = other.get_value_px()  # type: ignore
            if oth_val == 0:
                raise ZeroDivisionError
            return self.from_app_font(self.get_value_px() / oth_val)
        if hasattr(other, "get_value_mm100"):
            oth_val = other.get_value_mm100()  # type: ignore
            if oth_val == 0:
                raise ZeroDivisionError
            return self.from_mm100(self.get_value_mm100() / oth_val)  # type: ignore
        if isinstance(other, (int, float)):
            if other == 0:
                raise ZeroDivisionError
            return self.from_app_font(self.value / other)  # type: ignore
        return NotImplemented

    def __rtruediv__(self, other: object) -> Self:
        if isinstance(other, (int, float)):
            if self.value == 0:
                raise ZeroDivisionError
            return self.from_app_font(other / self.value)  # type: ignore
        return NotImplemented

    def __abs__(self) -> float:
        return abs(self.value)

    def __lt__(self, other: object) -> bool:
        if isinstance(other, UnitAppFontBase) and other.__class__.__name__ == self.__class__.__name__:
            return self.value < other.value
        if hasattr(other, "get_value_px"):
            oth_val = other.get_value_px()  # type: ignore
            return self.get_value_px() < oth_val
        if hasattr(other, "get_value_mm100"):
            return self.get_value_mm100() < other.get_value_mm100()  # type: ignore
        with contextlib.suppress(Exception):
            return self.value < float(other)  # type: ignore
        return False

    def __le__(self, other: object) -> bool:
        if isinstance(other, UnitAppFontBase) and other.__class__.__name__ == self.__class__.__name__:
            return True if self.almost_equal(other.value) else self.value < other.value
        if hasattr(other, "get_value_mm100"):
            return self.get_value_mm100() <= other.get_value_mm100()  # type: ignore
        with contextlib.suppress(Exception):
            oth_val = float(other)  # type: ignore
            return True if self.almost_equal(oth_val) else self.value < oth_val
        return False

    def __gt__(self, other: object) -> bool:
        if isinstance(other, UnitAppFontBase) and other.__class__.__name__ == self.__class__.__name__:
            return self.value > other.value
        if hasattr(other, "get_value_mm100"):
            return self.get_value_mm100() > other.get_value_mm100()  # type: ignore
        with contextlib.suppress(Exception):
            return self.value > float(other)  # type: ignore
        return False

    def __ge__(self, other: object) -> bool:
        if isinstance(other, UnitAppFontBase) and other.__class__.__name__ == self.__class__.__name__:
            return True if self.almost_equal(other.value) else self.value > other.value
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

    @abstractmethod
    def get_value_oth_unit(self) -> float:
        """
        Return the other value of the unit. If X then Y is returned.
        If Y then X is returned.
        """
        raise NotImplementedError

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
        return UnitConvert.convert(num=self.get_value_px(), frm=UnitLength.PX, to=UnitLength.MM)

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
        # ratio is usually less th 1.0
        # The num of pixels should be greater then then number of AppFont.
        return 0 if self.value == 0 else self.value / self._ratio

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

    def get_value_app_font(self, kind: Any = None) -> float:
        """
        Gets instance value in ``AppFont`` units.

        Returns:
            float: Value in ``AppFont`` units.
            kind (Any): Ignored in the context of ``AppFont`` units. It is used in other units.
        """
        return self.value

    @classmethod
    def from_mm(cls, value: float) -> Self:
        """
        Get instance from ``mm`` value.

        Args:
            value (int): ``mm`` value.

        Returns:
            Self:
        """
        # pylint: disable=protected-access
        # pylint: disable=unnecessary-dunder-call
        px = UnitConvert.convert(num=value, frm=UnitLength.MM, to=UnitLength.PX)
        inst = super(UnitAppFontBase, cls).__new__(cls)  # type: ignore
        inst._set_ratio()
        app_font_val = px * inst._ratio
        inst.__init__(app_font_val)
        return inst

    @classmethod
    def from_mm10(cls, value: float) -> Self:
        """
        Get instance from ``1/10th mm`` value.

        Args:
            value (int): ``1/10th mm`` value.

        Returns:
            Self:
        """
        # pylint: disable=protected-access
        # pylint: disable=unnecessary-dunder-call
        px = UnitConvert.convert(num=value, frm=UnitLength.MM10, to=UnitLength.PX)
        inst = super(UnitAppFontBase, cls).__new__(cls)  # type: ignore
        inst._set_ratio()
        app_font_val = px * inst._ratio
        inst.__init__(app_font_val)
        return inst

    @classmethod
    def from_mm100(cls, value: int) -> Self:
        """
        Get instance from ``1/100th mm`` value.

        Args:
            value (int): ``1/100th mm`` value.

        Returns:
            Self:
        """
        # pylint: disable=protected-access
        # pylint: disable=unnecessary-dunder-call
        px = UnitConvert.convert(num=value, frm=UnitLength.MM100, to=UnitLength.PX)
        inst = super(UnitAppFontBase, cls).__new__(cls)  # type: ignore
        inst._set_ratio()
        app_font_val = px * inst._ratio
        inst.__init__(app_font_val)
        return inst

    @classmethod
    def from_pt(cls, value: float) -> Self:
        """
        Get instance from ``pt`` (points) value.

        Args:
            value (float): ``pt`` value.

        Returns:
            Self:
        """
        # pylint: disable=protected-access
        # pylint: disable=unnecessary-dunder-call
        px = UnitConvert.convert(num=value, frm=UnitLength.PT, to=UnitLength.PX)
        inst = super(UnitAppFontBase, cls).__new__(cls)  # type: ignore
        inst._set_ratio()
        app_font_val = px * inst._ratio
        inst.__init__(app_font_val)
        return inst

    @classmethod
    def from_px(cls, value: float) -> Self:
        """
        Get instance from ``px`` (pixel) value.

        Args:
            value (float): ``px`` value.

        Returns:
            Self:
        """
        # pylint: disable=protected-access
        # pylint: disable=unnecessary-dunder-call
        # When converting pixels to app font the ratio is applied.
        # The ratio usually less then 1.0
        # The num of pixels should be greater then then number of pixels in the app font.
        # 10 pixels usually equals 5 or 6 app font.
        inst = super(UnitAppFontBase, cls).__new__(cls)  # type: ignore
        inst._set_ratio()
        app_font_val = value * inst._ratio
        inst.__init__(app_font_val)
        return inst

    @classmethod
    def from_inch(cls, value: float) -> Self:
        """
        Get instance from ``in`` (inch) value.

        Args:
            value (int): ``in`` value.

        Returns:
            Self:
        """
        # pylint: disable=protected-access
        # pylint: disable=unnecessary-dunder-call
        px = UnitConvert.convert(num=value, frm=UnitLength.IN, to=UnitLength.PX)
        inst = super(UnitAppFontBase, cls).__new__(cls)  # type: ignore
        inst._set_ratio()
        app_font_val = px * inst._ratio
        inst.__init__(app_font_val)
        return inst

    @classmethod
    def from_inch10(cls, value: float) -> Self:
        """
        Get instance from ``1/10th in`` (inch) value.

        Args:
            value (int): ``1/10th in`` value.

        Returns:
            Self:
        """
        # pylint: disable=protected-access
        # pylint: disable=unnecessary-dunder-call
        px = UnitConvert.convert(num=value, frm=UnitLength.IN10, to=UnitLength.PX)
        inst = super(UnitAppFontBase, cls).__new__(cls)  # type: ignore
        inst._set_ratio()
        app_font_val = px * inst._ratio
        inst.__init__(app_font_val)
        return inst

    @classmethod
    def from_inch100(cls, value: float) -> Self:
        """
        Get instance from ``1/100th in`` (inch) value.

        Args:
            value (int): ``1/100th in`` value.

        Returns:
            Self:
        """
        # pylint: disable=protected-access
        # pylint: disable=unnecessary-dunder-call
        px = UnitConvert.convert(num=value, frm=UnitLength.IN100, to=UnitLength.PX)
        inst = super(UnitAppFontBase, cls).__new__(cls)  # type: ignore
        inst._set_ratio()
        app_font_val = px * inst._ratio
        inst.__init__(app_font_val)
        return inst

    @classmethod
    def from_inch1000(cls, value: int) -> Self:
        """
        Get instance from ``1/1,000th in`` (inch) value.

        Args:
            value (int): ``1/1,000th in`` value.

        Returns:
            Self:
        """
        # pylint: disable=protected-access
        # pylint: disable=unnecessary-dunder-call
        px = UnitConvert.convert(num=value, frm=UnitLength.IN1000, to=UnitLength.PX)
        inst = super(UnitAppFontBase, cls).__new__(cls)  # type: ignore
        inst._set_ratio()
        app_font_val = px * inst._ratio
        inst.__init__(app_font_val)
        return inst

    @classmethod
    def from_cm(cls, value: float) -> Self:
        """
        Get instance from ``cm`` value.

        Args:
            value (int): ``cm`` value.

        Returns:
            Self:
        """
        # pylint: disable=protected-access
        # pylint: disable=unnecessary-dunder-call
        px = UnitConvert.convert(num=value, frm=UnitLength.CM, to=UnitLength.PX)
        inst = super(UnitAppFontBase, cls).__new__(cls)  # type: ignore
        inst._set_ratio()
        app_font_val = px * inst._ratio
        inst.__init__(app_font_val)
        return inst

    @classmethod
    def from_app_font(cls, value: float, kind: Any = None) -> Self:
        """
        Get instance from ``AppFont`` value.

        Args:
            value (int): ``AppFont`` value.
            kind (PointSizeKind, optional): The kind of ``AppFont`` to use.
                This is not used in the context of ``AppFont`` units.

        Returns:
            Self:
        """
        # pylint: disable=unnecessary-dunder-call
        return cls(value)

    @classmethod
    def from_unit_val(cls, value: UnitT | float | int) -> Self:
        """
        Get instance from ``UnitT`` or float value.

        Args:
            value (UnitT, float, int): ``UnitT`` or float value. If float then it is assumed to be in ``mm`` units.

        Returns:
            Self:
        """
        if isinstance(value, UnitAppFontBase) and value.__class__.__name__ == cls.__name__:
            return cls(value.value)

        try:
            unit_val = value.get_value_px()  # type: ignore
            return cls.from_px(unit_val)
        except AttributeError:
            return cls.from_app_font(float(value))  # type: ignore
