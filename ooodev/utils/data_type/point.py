from __future__ import annotations
from dataclasses import dataclass
from ooodev.utils.decorator import enforce


# Note that from __future__ import annotations converts annotations to string.
# this means that @enforce.enforce_types will see string as type. This is fine in
# most cases. Especially for built in types.
@enforce.enforce_types
@dataclass(frozen=True)
class Point:
    """Represents a X and Y values."""

    x: int
    y: int

    def __copy__(self) -> Point:
        return Point(self.x, self.y)

    def copy(self) -> Point:
        """
        Copies the instance

        Returns:
            Point: Copy of the instance

        .. versionadded:: 0.47.5
        """
        return self.__copy__()
