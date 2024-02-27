from __future__ import annotations
from dataclasses import dataclass
from ooodev.utils.validation import check
from ooodev.utils.decorator import enforce


# Note that from __future__ import annotations converts annotations to string.
# this means that @enforce.enforce_types will see string as type. This is fine in
# most cases. Especially for built in types.
@enforce.enforce_types
@dataclass(frozen=True)
class WidthHeightPercent:
    """Represents a Width and Height values from ``0`` to ``100``."""

    width: int
    height: int

    def __post_init__(self) -> None:
        check(
            self.width >= 0 and self.width <= 100,
            f"{self}",
            f"Width value of {self.width} is out of range. Value must be from 0 to 100.",
        )
        check(
            self.height >= 0 and self.height <= 100,
            f"{self}",
            f"Height value of {self.height} is out of range. Value must be from 0 to 100.",
        )
