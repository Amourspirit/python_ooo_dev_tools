from __future__ import annotations
from typing import TYPE_CHECKING
import contextlib
from dataclasses import dataclass
from ooodev.utils.data_type.base_int_value import BaseIntValue

if TYPE_CHECKING:
    from ooodev.units.angle_t import AngleT

# Note that from __future__ import annotations converts annotations to string.
# this means that @enforce.enforce_types will see string as type. This is fine in
# most cases. Especially for built in types.


def _to_positive_angle(angle: int) -> int:
    if not isinstance(angle, int):
        raise TypeError(f"Expected type int for angle. Got {type(angle).__name__}")
    # coverts all ints, negative or positive into a positive angle.
    # eg: -100 becomes 3500, 3800 becomes 200
    return (36000000 + angle) % 3600


@dataclass(unsafe_hash=True)
class Angle10(BaseIntValue):
    """
    Represents an angle value from ``0`` to ``3599``.

    All input integers are converted into a positive angle.

    Example:
        .. code-block::

            >>> print(Angle10(3600))
            Angle10(Value=0)
            >>> print(Angle10(4130))
            Angle10(Value=530)
            >>> print(Angle10(-2350))
            Angle10(Value=1250)
            >>> print(Angle10(17940))
            Angle10(Value=3540)

    .. versionadded:: 0.17.4
    """

    def __post_init__(self) -> None:
        self.value = _to_positive_angle(self.value)

    def _from_int(self, value: int) -> Angle10:
        return Angle10(_to_positive_angle(value))

    # region math and comparison

    def __eq__(self, other: object) -> bool:
        # for some reason BaseIntValue __eq__ is not picked up.
        # I suspect this is due to this class being a dataclass.
        if isinstance(other, Angle10):
            return self.value == other.value
        with contextlib.suppress(AttributeError):
            return self.get_angle100() == other.get_angle100()  # type: ignore
        with contextlib.suppress(AttributeError):
            return self.get_angle() == other.get_angle()  # type: ignore
        if isinstance(other, int):
            return self.value == other
        return False

    def __add__(self, other: object) -> Angle10:
        if hasattr(other, "get_angle10"):
            oth_val = other.get_angle10()  # type: ignore
            return self.from_angle10(self.value + oth_val)  # type: ignore

        if hasattr(other, "get_angle"):
            oth_val = other.get_angle() * 10  # type: ignore
            return self.from_angle10(self.value + oth_val)  # type: ignore

        if isinstance(other, (int, float)):
            return self.from_angle10(self.value + int(other))  # type: ignore

        return NotImplemented

    def __radd__(self, other: object) -> Angle10:
        return self if other == 0 else self.__add__(other)

    def __sub__(self, other: object) -> Angle10:
        if hasattr(other, "get_angle10"):
            oth_val = other.get_angle10()  # type: ignore
            return self.from_angle10(self.value - oth_val)  # type: ignore

        if hasattr(other, "get_angle"):
            oth_val = other.get_angle() * 10  # type: ignore
            return self.from_angle10(self.value - oth_val)  # type: ignore

        if isinstance(other, (int, float)):
            return self.from_angle10(self.value - int(other))  # type: ignore
        return NotImplemented

    def __rsub__(self, other: object) -> Angle10:
        if isinstance(other, (int, float)):
            self_val = self.get_angle10()
            return self.from_angle10(int(other) - self_val)  # type: ignore
        return NotImplemented

    def __mul__(self, other: object) -> Angle10:
        if hasattr(other, "get_angle10"):
            oth_val = other.get_angle10()  # type: ignore
            return self.from_angle10(self.value * oth_val)  # type: ignore

        if hasattr(other, "get_angle"):
            oth_val = other.get_angle() * 10  # type: ignore
            return self.from_angle10(self.value * oth_val)  # type: ignore

        if isinstance(other, (int, float)):
            return self.from_angle10(self.value * int(other))  # type: ignore

        return NotImplemented

    def __rmul__(self, other: int) -> Angle10:
        return self if other == 0 else self.__mul__(other)

    def __truediv__(self, other: object) -> Angle10:
        if hasattr(other, "get_angle10"):
            oth_val = other.get_angle10()  # type: ignore
            if oth_val == 0:
                raise ZeroDivisionError
            return self.from_angle10(self.value // oth_val)  # type: ignore

        if hasattr(other, "get_angle"):
            oth_val = other.get_angle() * 10  # type: ignore
            if oth_val == 0:
                raise ZeroDivisionError
            return self.from_angle10(self.value // oth_val)  # type: ignore
        if isinstance(other, (int, float)):
            if other == 0:
                raise ZeroDivisionError
            return self.from_angle10(self.value // other)  # type: ignore
        return NotImplemented

    def __rtruediv__(self, other: object) -> Angle10:
        if isinstance(other, (int, float)):
            if self.value == 0:
                raise ZeroDivisionError
            return self.from_angle10(other // self.value)  # type: ignore
        return NotImplemented

    def __abs__(self) -> int:
        return abs(self.value)

    def __lt__(self, other: object) -> bool:
        if hasattr(other, "get_angle10"):
            return self.value < other.get_angle10()  # type: ignore
        if hasattr(other, "get_angle"):
            return self.value < (other.get_angle() * 10)  # type: ignore
        with contextlib.suppress(Exception):
            return self.value < int(other)  # type: ignore
        return False

    def __le__(self, other: object) -> bool:
        if hasattr(other, "get_angle10"):
            return self.value <= other.get_angle10()  # type: ignore
        if hasattr(other, "get_angle"):
            return self.value <= (other.get_angle() * 10)  # type: ignore
        with contextlib.suppress(Exception):
            return self.value <= int(other)  # type: ignore
        return False

    def __gt__(self, other: object) -> bool:
        if hasattr(other, "get_angle10"):
            return self.value > other.get_angle10()  # type: ignore
        if hasattr(other, "get_angle"):
            return self.value > (other.get_angle() * 10)  # type: ignore
        with contextlib.suppress(Exception):
            return self.value > int(other)  # type: ignore
        return False

    def __ge__(self, other: object) -> bool:
        if hasattr(other, "get_angle10"):
            return self.value >= other.get_angle10()  # type: ignore
        if hasattr(other, "get_angle"):
            return self.value >= (other.get_angle() * 10)  # type: ignore
        with contextlib.suppress(Exception):
            return self.value >= int(other)  # type: ignore
        return False

    # endregion math and comparison

    def get_angle(self) -> int:
        """Gets Angle10 Value as ``degrees``"""
        if self.value == 0:
            return 0
        return round(self.value / 10)

    def get_angle10(self) -> int:
        """Gets Angle10 Value as ``1/10 degree``"""
        return self.value

    def get_angle100(self) -> int:
        """Gets Angle10 Value as ``1/100 degree``"""
        return self.value * 10

    @staticmethod
    def from_angle(value: int) -> Angle10:
        """
        Get an angle from ``degree`` units.

        Args:
            value (int): Angle10 in ``degree`` units.

        Returns:
            Angle10:
        """
        return Angle10(0) if value == 0 else Angle10(value * 10)

    @staticmethod
    def from_angle10(value: int) -> Angle10:
        """
        Get an angle from ``1/10 degree`` units.

        Args:
            value (int): Angle10 in ``1/10 degree`` units.

        Returns:
            Angle10:
        """
        return Angle10(value)

    @staticmethod
    def from_angle100(value: int) -> Angle10:
        """
        Get an angle from ``1/100 degree`` units.

        Args:
            value (int): Angle10 in ``1/10 degree`` units.

        Returns:
            Angle10:
        """
        return Angle10(0) if value == 0 else Angle10(round(value / 10))

    @classmethod
    def from_unit_val(cls, value: AngleT | int) -> Angle10:
        """
        Get instance from ``Angle10`` or int value.

        Args:
            value (Angle10, int): ``Angle10`` or int value. If int then it is assumed to be in ``1/10th`` degrees.

        Returns:
            Angle10:

        .. versionadded:: 0.32.0
        """
        try:
            unit_10 = value.get_angle10()  # type: ignore
            return cls.from_angle10(unit_10)
        except AttributeError:
            return cls.from_angle10(int(value))  # type: ignore
