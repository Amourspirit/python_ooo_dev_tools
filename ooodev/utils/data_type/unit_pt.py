from __future__ import annotations
from typing import TypeVar, Type
from dataclasses import dataclass
from ..validation import check
from ..decorator import enforce
from .base_int_value import BaseIntValue
from ..unit_convert import UnitConvert, Length

_TUnitPT = TypeVar(name="_TUnitPT", bound="UnitPT")

# Note that from __future__ import annotations converts annotations to string.
# this means that @enforce.enforce_types will see string as type. This is fine in
# most cases. Especially for built in types.
@enforce.enforce_types
@dataclass(unsafe_hash=True)
class UnitPT(BaseIntValue):
    """Represents a ``PT`` (points) value."""

    def __post_init__(self) -> None:
        check(
            self.value >= 0,
            f"{self}",
            f"Value of {self.value} is out of range. Value must be a positive number.",
        )

    def _from_int(self, value: int) -> _TUnitPT:
        inst = super(UnitPT, self.__class__).__new__(self.__class__)
        return inst.__init__(value)

    def __eq__(self, other: object) -> bool:
        # for some reason BaseIntValue __eq__ is not picked up.
        # I suspect this is due to this class being a dataclass.
        try:
            i = int(other)
            return i == self.value
        except Exception as e:
            return False

    def get_value_mm(self) -> float:
        """
        Gets instance value converted to Size in ``mm`` units.

        Returns:
            int: Value in ``mm`` units.
        """
        return float(UnitConvert.convert(num=self.value, frm=Length.PT, to=Length.MM))

    def get_value_mm_100(self) -> int:
        """
        Gets instance value converted to Size in ``1/100th mm`` units.

        Returns:
            int: Value in ``1/100th mm`` units.
        """
        return round(UnitConvert.convert(num=self.value, frm=Length.PT, to=Length.MM100))

    @classmethod
    def from_mm(cls: Type[_TUnitPT], value: float) -> _TUnitPT:
        """
        Get instance from ``mm`` value.

        Args:
            value (int): ``mm`` value.

        Returns:
            UnitPT:
        """
        inst = super(UnitPT, cls).__new__(cls)
        return inst.__init__(round(UnitConvert.convert(num=value, frm=Length.MM, to=Length.PT)))

    @classmethod
    def from_mm_100(cls: Type[_TUnitPT], value: int) -> _TUnitPT:
        """
        Get instance from ``1/100th mm`` value.

        Args:
            value (int): ``1/100th mm`` value.

        Returns:
            UnitPT:
        """
        inst = super(UnitPT, cls).__new__(cls)
        return inst.__init__(round(UnitConvert.convert(num=value, frm=Length.MM100, to=Length.PT)))
