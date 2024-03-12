from __future__ import annotations
from typing import TYPE_CHECKING
import contextlib
from dataclasses import dataclass
from ooodev.utils.data_type.base_int_value import BaseIntValue

if TYPE_CHECKING:
    from ooodev.units.angle_t import AngleT


def _to_positive_angle(angle: int) -> int:
    if not isinstance(angle, int):
        raise TypeError(f"Expected type int for angle. Got {type(angle).__name__}")
    # coverts all ints, negative or positive into a positive angle.
    # eg: -10 becomes 35000, 38000 becomes 2000
    return (360000000 + angle) % 36000


@dataclass(unsafe_hash=True)
class Angle100(BaseIntValue):
    """
    Represents an angle value from ``0`` to ``35999`` in ``1/100 units``.

    All input integers are converted into a positive angle.

    Example:
        .. code-block::

            >>> print(Angle100(36000))
            Angle100(Value=0)
            >>> print(Angle(41300))
            Angle100(Value=5300)
            >>> print(Angle(-23500))
            Angle(Value=12500)
            >>> Angle100(Angle(179400))
            Angle(Value=35400)

    .. versionadded:: 0.17.4
    """

    def __post_init__(self) -> None:
        self.value = _to_positive_angle(self.value)

    def _from_int(self, value: int) -> Angle100:
        return Angle100(_to_positive_angle(value))

    # region math and comparison

    def __eq__(self, other: object) -> bool:
        # for some reason BaseIntValue __eq__ is not picked up.
        # I suspect this is due to this class being a dataclass.
        if isinstance(other, Angle100):
            return self.value == other.value
        with contextlib.suppress(AttributeError):
            return self.get_angle100() == other.get_angle100()  # type: ignore
        with contextlib.suppress(AttributeError):
            return self.get_angle() == other.get_angle()  # type: ignore
        if isinstance(other, int):
            return self.value == other
        return False

    def __add__(self, other: object) -> Angle100:
        if hasattr(other, "get_angle100"):
            oth_val = other.get_angle100()  # type: ignore
            return self.from_angle100(self.value + oth_val)  # type: ignore

        if hasattr(other, "get_angle"):
            oth_val = other.get_angle() * 100  # type: ignore
            return self.from_angle100(self.value + oth_val)  # type: ignore

        if isinstance(other, (int, float)):
            return self.from_angle100(self.value + int(other))  # type: ignore

        return NotImplemented

    def __radd__(self, other: object) -> Angle100:
        return self if other == 0 else self.__add__(other)

    def __sub__(self, other: object) -> Angle100:
        if hasattr(other, "get_angle100"):
            oth_val = other.get_angle100()  # type: ignore
            return self.from_angle100(self.value - oth_val)  # type: ignore

        if hasattr(other, "get_angle"):
            oth_val = other.get_angle() * 100  # type: ignore
            return self.from_angle100(self.value - oth_val)  # type: ignore

        if isinstance(other, (int, float)):
            return self.from_angle100(self.value - int(other))  # type: ignore
        return NotImplemented

    def __rsub__(self, other: object) -> Angle100:
        if isinstance(other, (int, float)):
            self_val = self.get_angle100()
            return self.from_angle100(int(other) - self_val)  # type: ignore
        return NotImplemented

    def __mul__(self, other: object) -> Angle100:
        if hasattr(other, "get_angle100"):
            oth_val = other.get_angle100()  # type: ignore
            return self.from_angle100(self.value * oth_val)  # type: ignore

        if hasattr(other, "get_angle"):
            oth_val = other.get_angle() * 100  # type: ignore
            return self.from_angle100(self.value * oth_val)  # type: ignore

        if isinstance(other, (int, float)):
            return self.from_angle100(self.value * int(other))  # type: ignore

        return NotImplemented

    def __rmul__(self, other: int) -> Angle100:
        return self if other == 0 else self.__mul__(other)

    def __truediv__(self, other: object) -> Angle100:
        if hasattr(other, "get_angle100"):
            oth_val = other.get_angle100()  # type: ignore
            if oth_val == 0:
                raise ZeroDivisionError
            return self.from_angle100(self.value // oth_val)  # type: ignore

        if hasattr(other, "get_angle"):
            oth_val = other.get_angle() * 100  # type: ignore
            if oth_val == 0:
                raise ZeroDivisionError
            return self.from_angle100(self.value // oth_val)  # type: ignore
        if isinstance(other, (int, float)):
            if other == 0:
                raise ZeroDivisionError
            return self.from_angle100(self.value // other)  # type: ignore
        return NotImplemented

    def __rtruediv__(self, other: object) -> Angle100:
        if isinstance(other, (int, float)):
            if self.value == 0:
                raise ZeroDivisionError
            return self.from_angle100(other // self.value)  # type: ignore
        return NotImplemented

    def __abs__(self) -> int:
        return abs(self.value)

    def __lt__(self, other: object) -> bool:
        if hasattr(other, "get_angle100"):
            return self.value < other.get_angle100()  # type: ignore
        if hasattr(other, "get_angle"):
            return self.value < (other.get_angle() * 100)  # type: ignore
        with contextlib.suppress(Exception):
            return self.value < int(other)  # type: ignore
        return False

    def __le__(self, other: object) -> bool:
        if hasattr(other, "get_angle100"):
            return self.value <= other.get_angle100()  # type: ignore
        if hasattr(other, "get_angle"):
            return self.value <= (other.get_angle() * 100)  # type: ignore
        with contextlib.suppress(Exception):
            return self.value <= int(other)  # type: ignore
        return False

    def __gt__(self, other: object) -> bool:
        if hasattr(other, "get_angle100"):
            return self.value > other.get_angle100()  # type: ignore
        if hasattr(other, "get_angle"):
            return self.value > (other.get_angle() * 100)  # type: ignore
        with contextlib.suppress(Exception):
            return self.value > int(other)  # type: ignore
        return False

    def __ge__(self, other: object) -> bool:
        if hasattr(other, "get_angle100"):
            return self.value >= other.get_angle100()  # type: ignore
        if hasattr(other, "get_angle"):
            return self.value >= (other.get_angle() * 100)  # type: ignore
        with contextlib.suppress(Exception):
            return self.value >= int(other)  # type: ignore
        return False

    # endregion math and comparison

    def get_angle(self) -> int:
        """Gets Angle Value as ``degrees``"""
        if self.value == 0:
            return 0
        return round(self.value / 100)

    def get_angle10(self) -> int:
        """Gets Angle Value as ``1/10 degree``"""
        if self.value == 0:
            return 0
        return round(self.value / 10)

    def get_angle100(self) -> int:
        """Gets Angle Value as ``1/100 degree``"""
        return self.value

    @staticmethod
    def from_angle(value: int) -> Angle100:
        """
        Get an angle from ``degree`` units.

        Args:
            value (int): Angle in ``degree`` units.

        Returns:
            Angle100:
        """
        return Angle100(0) if value == 0 else Angle100(value * 100)

    @staticmethod
    def from_angle10(value: int) -> Angle100:
        """
        Get an angle from ``1/10 degree`` units.

        Args:
            value (int): Angle in ``1/10 degree`` units.

        Returns:
            Angle:
        """
        return Angle100(0) if value == 0 else Angle100(round(value / 10))

    @staticmethod
    def from_angle100(value: int) -> Angle100:
        """
        Get an angle from ``1/100 degree`` units.

        Args:
            value (int): Angle in ``1/100 degree`` units.

        Returns:
            Angle:
        """
        return Angle100(0) if value == 0 else Angle100(round(value))

    @classmethod
    def from_unit_val(cls, value: AngleT | int) -> Angle100:
        """
        Get instance from ``Angle100`` or int value.

        Args:
            value (Angle100, int): ``Angle100`` or int value. If int then it is assumed to be in ``1/100th`` degrees.

        Returns:
            Angle100:

        .. versionadded:: 0.32.0
        """
        try:
            unit_100 = value.get_angle100()  # type: ignore
            return cls.from_angle100(unit_100)
        except AttributeError:
            return cls.from_angle100(int(value))  # type: ignore
