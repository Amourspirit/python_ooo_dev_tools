from __future__ import annotations
from abc import abstractmethod
from typing import TypeVar
from dataclasses import dataclass
import math


# Note that from __future__ import annotations converts annotations to string.
# this means that @enforce.enforce_types will see string as type. This is fine in
# most cases. Especially for built in types.

_TBaseFloatValue = TypeVar("_TBaseFloatValue", bound="BaseFloatValue")


@dataclass(unsafe_hash=True)
class BaseFloatValue:
    """Base class for float Value"""

    value: float
    """Float value."""

    @abstractmethod
    def _from_float(self: _TBaseFloatValue, value: float) -> _TBaseFloatValue:
        ...

    def almost_equal(self, val: float, epsilon: float = 1e-9) -> bool:
        """
        Comparing float values directly using equality (``==``) can sometimes lead to
        unexpected results due to the way floating-point numbers are represented in computers.
        A small rounding error can make two floats that should be equal appear unequal.

        A common way to compare floats is to check if the absolute difference between them
        is less than a small number, often called the machine epsilon.

        In this function, ``epsilon`` is the maximum difference for which ``a`` and ``b``
        are considered equal. You can adjust ``epsilon`` based on the precision you need.

        Args:
            val (float): The value to compare with.
            epsilon (float): The maximum difference for which ``a`` and ``b`` are considered equal.

        Returns:
            bool: True if current value and ``val`` are considered equal, False otherwise.
        """
        return abs(self.value - val) < epsilon

    # Override "int" method
    def __float__(self) -> float:
        return self.value

    def __int__(self) -> float:
        return int(self.value)

    def __add__(self: _TBaseFloatValue, other: object) -> _TBaseFloatValue:
        try:
            i = float(other)  # type: ignore
            return self._from_float(self.value + i)
        except AssertionError:
            raise
        except Exception:
            return NotImplemented

    def __radd__(self: _TBaseFloatValue, other: object) -> _TBaseFloatValue:
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

    def __sub__(self: _TBaseFloatValue, other: object) -> _TBaseFloatValue:
        try:
            i = float(other)  # type: ignore
            return self._from_float(self.value - i)
        except AssertionError:
            raise
        except Exception:
            return NotImplemented

    def __rsub__(self: _TBaseFloatValue, other: object) -> _TBaseFloatValue:
        try:
            i = float(other)  # type: ignore
            return self._from_float(i - self.value)
        except AssertionError:
            raise
        except Exception:
            return NotImplemented

    def __mul__(self: _TBaseFloatValue, other: object) -> _TBaseFloatValue:
        try:
            i = float(other)  # type: ignore
            return self._from_float(self.value * i)
        except AssertionError:
            raise
        except Exception:
            return NotImplemented

    def __rmul__(self: _TBaseFloatValue, other: int) -> _TBaseFloatValue:
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
