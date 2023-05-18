from __future__ import annotations
from abc import abstractmethod
from typing import TYPE_CHECKING
from dataclasses import dataclass
import math

if TYPE_CHECKING:
    try:
        from typing import Self
    except ImportError:
        from typing_extensions import Self

# Note that from __future__ import annotations converts annotations to string.
# this means that @enforce.enforce_types will see string as type. This is fine in
# most cases. Especially for built in types.


@dataclass(unsafe_hash=True)
class BaseFloatValue:
    """Base class for float Value"""

    value: float
    """Float value."""

    @abstractmethod
    def _from_float(self, value: float) -> Self:
        ...

    # Override "int" method
    def __float__(self) -> float:
        return self.value

    def __int__(self) -> float:
        return int(self.value)

    def __add__(self, other: object) -> Self:
        try:
            i = float(other)  # type: ignore
            return self._from_float(self.value + i)
        except AssertionError:
            raise
        except Exception:
            return NotImplemented

    def __radd__(self, other: object) -> Self:
        # angle = sum([ang1, ang2, ang3])
        # will result in TypeError because sum() start with 0
        # this will force a call to __radd__
        return self if other == 0 else self.__add__(other)

    def __eq__(self, other: object) -> bool:
        # By default, __ne__() delegates to __eq__() and inverts the result unless it is NotImplemented.
        # There are no other implied relationships among the comparison operators,
        # for example, the truth of (x<y or x==y) does not imply x<=y.
        try:
            i = float(other)  # type: ignore
            return math.isclose(i, self.value)
        except Exception as e:
            return False

    def __sub__(self, other: object) -> Self:
        try:
            i = float(other)  # type: ignore
            return self._from_float(self.value - i)
        except AssertionError:
            raise
        except Exception:
            return NotImplemented

    def __rsub__(self, other: object) -> Self:
        try:
            i = float(other)  # type: ignore
            return self._from_float(i - self.value)
        except AssertionError:
            raise
        except Exception:
            return NotImplemented

    def __mul__(self, other: object) -> Self:
        try:
            i = float(other)  # type: ignore
            return self._from_float(self.value * i)
        except AssertionError:
            raise
        except Exception:
            return NotImplemented

    def __rmul__(self, other: int) -> Self:
        return self if other == 0 else self.__mul__(other)

    def __lt__(self, other: object) -> bool:
        try:
            i = float(other)  # type: ignore
            return False if math.isclose(i, self.value) else self.value < i
        except Exception:
            return NotImplemented

    def __le__(self, other: object) -> bool:
        try:
            i = float(other)  # type: ignore
            return True if math.isclose(i, self.value) else self.value <= i
        except Exception:
            return NotImplemented

    def __gt__(self, other: object) -> bool:
        try:
            i = float(other)  # type: ignore
            return False if math.isclose(i, self.value) else self.value > i
        except Exception:
            return NotImplemented

    def __ge__(self, other: object) -> bool:
        try:
            i = float(other)  # type: ignore
            return True if math.isclose(i, self.value) else self.value >= i
        except Exception:
            return NotImplemented

    def __abs__(self) -> int:
        return abs(round(self.value))
