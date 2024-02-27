from __future__ import annotations
from dataclasses import dataclass
from ooodev.utils.validation import check
from ooodev.utils.decorator import enforce
from ooodev.utils.data_type.base_int_value import BaseIntValue


# Note that from __future__ import annotations converts annotations to string.
# this means that @enforce.enforce_types will see string as type. This is fine in
# most cases. Especially for built in types.


@enforce.enforce_types
@dataclass(unsafe_hash=True)
class PolySides(BaseIntValue):
    """Represents Polygon Sides value from ``3`` to ``30``."""

    def __post_init__(self) -> None:
        check(
            self.value >= 3 and self.value <= 30,
            f"{self}",
            f"Value of {self.value} is out of range. Value must be from 3 to 30.",
        )

    def _from_int(self, int) -> PolySides:
        return PolySides(int)

    def __eq__(self, other: object) -> bool:
        # for some reason BaseIntValue __eq__ is not picked up.
        # I suspect this is due to this class being a dataclass.
        try:
            i = int(other)  # type: ignore
            return i == self.value
        except Exception as e:
            return False
