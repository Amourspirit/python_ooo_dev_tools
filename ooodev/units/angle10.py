from __future__ import annotations
import contextlib
from dataclasses import dataclass
from ooodev.utils.data_type.base_int_value import BaseIntValue

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

    def __eq__(self, other: object) -> bool:
        # for some reason BaseIntValue __eq__ is not picked up.
        # I suspect this is due to this class being a dataclass.
        if isinstance(other, Angle10):
            return self.value == other.value
        with contextlib.suppress(AttributeError):
            return self.get_angle100() == other.get_angle100()  # type: ignore
        return False

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
