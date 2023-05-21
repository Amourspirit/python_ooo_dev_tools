from __future__ import annotations
from abc import abstractmethod
from typing import TypeVar
from dataclasses import dataclass


# Note that from __future__ import annotations converts annotations to string.
# this means that @enforce.enforce_types will see string as type. This is fine in
# most cases. Especially for built in types.

_BaseIntValue = TypeVar("_BaseIntValue", bound="BaseIntValue")


@dataclass(unsafe_hash=True)
class BaseIntValue:
    """Base class for Int Value"""

    value: int
    """Int value."""

    @abstractmethod
    def _from_int(self: _BaseIntValue, value: int) -> _BaseIntValue:
        ...

    # Override "int" method
    def __int__(self) -> int:
        return self.value

    def __add__(self: _BaseIntValue, other: object) -> _BaseIntValue:
        try:
            i = int(other)  # type: ignore
            return self._from_int(self.value + i)
        except AssertionError:
            raise
        except Exception:
            return NotImplemented

    def __radd__(self: _BaseIntValue, other: object) -> _BaseIntValue:
        # angle = sum([ang1, ang2, ang3])
        # will result in TypeError because sum() start with 0
        # this will force a call to __radd__
        return self if other == 0 else self.__add__(other)

    def __eq__(self, other: object) -> bool:
        # By default, __ne__() delegates to __eq__() and inverts the result unless it is NotImplemented.
        # There are no other implied relationships among the comparison operators,
        # for example, the truth of (x<y or x==y) does not imply x<=y.
        try:
            i = int(other)  # type: ignore
            return i == self.value
        except Exception as e:
            return False

    def __sub__(self: _BaseIntValue, other: object) -> _BaseIntValue:
        try:
            i = int(other)  # type: ignore
            return self._from_int(self.value - i)
        except AssertionError:
            raise
        except Exception:
            return NotImplemented

    def __rsub__(self: _BaseIntValue, other: object) -> _BaseIntValue:
        try:
            i = int(other)  # type: ignore
            return self._from_int(i - self.value)
        except AssertionError:
            raise
        except Exception:
            return NotImplemented

    def __mul__(self: _BaseIntValue, other: object) -> _BaseIntValue:
        try:
            i = int(other)  # type: ignore
            return self._from_int(self.value * i)
        except AssertionError:
            raise
        except Exception:
            return NotImplemented

    def __rmul__(self: _BaseIntValue, other: int) -> _BaseIntValue:
        return self if other == 0 else self.__mul__(other)

    def __lt__(self, other: object) -> bool:
        try:
            i = int(other)  # type: ignore
            return self.value < i
        except Exception:
            return NotImplemented

    def __le__(self, other: object) -> bool:
        try:
            i = int(other)  # type: ignore
            return self.value <= i
        except Exception:
            return NotImplemented

    def __gt__(self, other: object) -> bool:
        try:
            i = int(other)  # type: ignore
            return self.value > i
        except Exception:
            return NotImplemented

    def __ge__(self, other: object) -> bool:
        try:
            i = int(other)  # type: ignore
            return self.value >= i
        except Exception:
            return NotImplemented

    def __abs__(self) -> int:
        return abs(self.value)

    def __truediv__(self: _BaseIntValue, other) -> _BaseIntValue:
        try:
            i = int(other)
            return self._from_int(round(self.value / i))
        except Exception:
            return NotImplemented

    def __rtruediv__(self: _BaseIntValue, other) -> _BaseIntValue:
        try:
            i = int(other)
            return self._from_int(round(i / self.value))
        except Exception:
            return NotImplemented
