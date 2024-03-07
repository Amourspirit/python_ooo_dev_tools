from __future__ import annotations
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from typing_extensions import Protocol
else:
    Protocol = object


class AngleT(Protocol):
    """
    Represents an angle value from ``0`` to ``359`` in ``degrees``.

    .. versionadded:: 0.32.0
    """

    def get_angle(self) -> int:
        """Gets Angle Value as ``degrees``"""
        ...

    def get_angle10(self) -> int:
        """Gets Angle Value as ``1/10 degree``"""
        ...

    def get_angle100(self) -> int:
        """Gets Angle Value as ``1/100 degree``"""
        ...

    @property
    def value(self) -> int:
        """Gets/Sets - Actual value"""
        ...

    @value.setter
    def value(self, value: int) -> None: ...
