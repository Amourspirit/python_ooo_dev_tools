from __future__ import annotations
from typing import TYPE_CHECKING
from dataclasses import dataclass
from ..validation import check
from ..decorator import enforce
from .base_int_value import BaseIntValue
from ..unit_convert import UnitConvert

if TYPE_CHECKING:
    try:
        from typing import Self
    except ImportError:
        from typing_extensions import Self

# Note that from __future__ import annotations converts annotations to string.
# this means that @enforce.enforce_types will see string as type. This is fine in
# most cases. Especially for built in types.
@enforce.enforce_types
@dataclass(unsafe_hash=True)
class Unit100MM(BaseIntValue):
    """Represents a intensity value from ``0`` to ``100``."""

    def __post_init__(self) -> None:
        check(
            self.value >= 0,
            f"{self}",
            f"Value of {self.value} is out of range. Value must be a positive number.",
        )

    def _from_int(self, int) -> Self:
        return Unit100MM(int)

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
        return UnitConvert.convert_mm100_mm(self.value)

    def from_mm(value: float) -> Unit100MM:
        """
        Get instance from ``mm`` value.

        Args:
            value (int): ``mm`` value.

        Returns:
            Unit100MM:
        """
        return Unit100MM(UnitConvert.convert_mm_mm100(value))
