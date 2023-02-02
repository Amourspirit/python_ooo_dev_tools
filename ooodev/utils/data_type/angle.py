from __future__ import annotations
from typing import TYPE_CHECKING
from dataclasses import dataclass
from .base_int_value import BaseIntValue

if TYPE_CHECKING:
    try:
        from typing import Self
    except ImportError:
        from typing_extensions import Self

# Note that from __future__ import annotations converts annotations to string.
# this means that @enforce.enforce_types will see string as type. This is fine in
# most cases. Especially for built in types.


def _to_positive_angle(angle: int) -> int:
    if not isinstance(angle, int):
        raise TypeError(f"Expected type int for angle. Got {type(angle).__name__}")
    # coverts all ints, negative or positive into a positive angel.
    # eg: -10 becomes 350, 380 becomes 20
    return (3600000 + angle) % 360


@dataclass(unsafe_hash=True)
class Angle(BaseIntValue):
    """
    Represents a angle value from ``0`` to ``359``.

    All input integers are converted into a positive angle.

    Example:
        .. code-block::

            >>> print(Angle(360))
            Angle(Value=0)
            >>> print(Angle(413))
            Angle(Value=53)
            >>> print(Angle(-235))
            Angle(Value=125)
            >>> print(Angle(1794))
            Angle(Value=354)

    .. versionchanged:: 0.8.1

        Now will accept any integer value, negative or positive.
    """

    def __post_init__(self) -> None:
        self.value = _to_positive_angle(self.value)

    def _from_int(self, value: int) -> Self:
        return Angle(_to_positive_angle(value))

    def __eq__(self, other: object) -> bool:
        # for some reason BaseIntValue __eq__ is not picked up.
        # I suspect this is due to this class being a dataclass.
        try:
            i = int(other)
            return i == self.value
        except Exception as e:
            return False
