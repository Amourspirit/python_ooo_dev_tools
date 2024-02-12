from __future__ import annotations
import contextlib
from typing import TypeVar, Type, TYPE_CHECKING
from dataclasses import dataclass

from ..utils.decorator import enforce
from .unit_convert import UnitConvert, UnitLength

if TYPE_CHECKING:
    from ooodev.units import UnitT

_TUnitMM100 = TypeVar(name="_TUnitMM100", bound="UnitMM100")


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

    @classmethod
    def from_mm(cls: Type[_TUnitMM100], value: float) -> _TUnitMM100:
        """
        Get instance from ``mm`` value.

        Args:
            value (int): ``mm`` value.

        Returns:
            UnitMM100:
        """
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
        inst = super(UnitMM100, cls).__new__(cls)  # type: ignore
        inst.__init__(round(UnitConvert.convert(num=value, frm=UnitLength.CM, to=UnitLength.MM100)))
        return inst

    @classmethod
    def from_unit_val(cls: Type[_TUnitMM100], value: UnitT | float) -> _TUnitMM100:
        """
        Get instance from ``UnitT`` or float value.

        Args:
            value (UnitT, float): ``UnitT`` or float value. If float then it is assumed to be in ``1/100th mm`` units.

        Returns:
            UnitMM100:
        """
        try:
            unit_100 = value.get_value_mm100()  # type: ignore
            return cls.from_mm100(unit_100)
        except AttributeError:
            return cls.from_mm100(round(value))  # type: ignore
