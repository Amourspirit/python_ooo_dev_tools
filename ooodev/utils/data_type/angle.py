from __future__ import annotations
from typing import TYPE_CHECKING
from dataclasses import dataclass
from ..validation import check
from ..decorator import enforce
from .base_int_value import BaseIntValue

if TYPE_CHECKING:
    try:
        from typing import Self
    except ImportError:
        from typing_extensions import Self

# Note that from __future__ import annotations converts annotations to string.
# this means that @enforce.enforce_types will see string as type. This is fine in
# most cases. Especially for built in types.


@enforce.enforce_types
@dataclass(frozen=True)
class Angle(BaseIntValue):
    """Represents a angle value from ``0`` to ``360``."""

    def __post_init__(self) -> None:
        check(
            self.Value >= 0 and self.Value <= 360,
            f"{self}",
            f"Value of {self.Value} is out of range. Value must be from 0 to 360.",
        )

    def _from_int(self, int) -> Self:
        return Angle(int)
