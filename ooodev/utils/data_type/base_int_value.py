from __future__ import annotations
from abc import abstractmethod
from typing import TYPE_CHECKING
from dataclasses import dataclass

if TYPE_CHECKING:
    try:
        from typing import Self
    except ImportError:
        from typing_extensions import Self

# Note that from __future__ import annotations converts annotations to string.
# this means that @enforce.enforce_types will see string as type. This is fine in
# most cases. Especially for built in types.


@dataclass(frozen=True)
class BaseIntValue:
    """Base class for Int Value"""

    Value: int
    """Int value."""

    @abstractmethod
    def _from_int(self, value: int) -> Self:
        ...

    # Override "int" method
    def __int__(self) -> int:
        return self.Value

    def __add__(self, other: object) -> Self:
        try:
            i = int(other)
            return self._from_int(self.Value + i)
        except AssertionError:
            raise
        except Exception:
            return NotImplemented

    def __radd__(self, other: object) -> Self:
        # angle = sum([ang1, ang2, ang3])
        # will result in TypeError becuase sum() start with 0
        # this will force a call to __radd__
        if other == 0:
            return self
        else:
            return self.__add__(other)

    def __eq__(self, other: object) -> bool:
        # By default, __ne__() delegates to __eq__() and inverts the result unless it is NotImplemented.
        # There are no other implied relationships among the comparison operators,
        # for example, the truth of (x<y or x==y) does not imply x<=y.
        try:
            i = int(other)
            return i == self.Value
        except Exception as e:
            return False

    def __sub__(self, other: object) -> Self:
        try:
            i = int(other)
            return self._from_int(self.Value - i)
        except AssertionError:
            raise
        except Exception:
            return NotImplemented

    def __rsub__(self, other: object) -> Self:
        try:
            i = int(other)
            return self._from_int(i=self.Value)
        except AssertionError:
            raise
        except Exception:
            return NotImplemented

    def __mul__(self, other: object) -> Self:
        try:
            i = int(other)
            return self._from_int(self.Value * i)
        except AssertionError:
            raise
        except Exception:
            return NotImplemented

    def __rmul__(self, other: int) -> Self:
        if other == 0:
            return self
        else:
            return self.__mul__(other)

    def __lt__(self, other: object) -> bool:
        try:
            i = int(other)
            return self.Value < i
        except Exception:
            return NotImplemented

    def __le__(self, other: object) -> bool:
        try:
            i = int(other)
            return self.Value <= i
        except Exception:
            return NotImplemented

    def __gt__(self, other: object) -> bool:
        try:
            i = int(other)
            return self.Value > i
        except Exception:
            return NotImplemented

    def __ge__(self, other: object) -> bool:
        try:
            i = int(other)
            return self.Value >= i
        except Exception:
            return NotImplemented

    def __abs__(self) -> int:
        return abs(self.Value)

    def __truediv__(self, other):
        try:
            i = int(other)
            return self._from_int(round(self.Value / i))
        except Exception:
            return NotImplemented

    def __rtruediv__(self, other):
        try:
            i = int(other)
            return self._from_int(round(i / self.Value))
        except Exception:
            return NotImplemented
