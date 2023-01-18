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
@dataclass(unsafe_hash=True)
class Byte(BaseIntValue):
    """Represents a intensity value from ``0`` to ``100``."""

    def __post_init__(self) -> None:
        check(
            self.value >= 0 and self.value <= 255,
            f"{self}",
            f"Value of {self.value} is out of range. Value must be from 0 to 255.",
        )

    def _from_int(self, int) -> Self:
        return Byte(int)