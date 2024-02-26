from __future__ import annotations
from typing import TYPE_CHECKING
from ooodev.format.inner.kind.format_kind import FormatKind as FormatKind

if TYPE_CHECKING:
    try:
        from typing import Protocol
    except ImportError:
        from typing_extensions import Protocol
else:
    Protocol = object


class AngleUnitT(Protocol):
    """
    Protocol Class for Angle units.

    .. seealso::

        :ref:`ns_units`
    """

    @property
    def value(self) -> float | int:
        """Angle actual value. Generally a ``float`` or ``int``"""
        ...

    def get_angle(self) -> int:
        """Gets Angle Value as ``degrees``"""
        ...

    def get_angle10(self) -> int:
        """Gets Angle Value as ``1/10 degree``"""
        ...

    def get_angle100(self) -> int:
        """Gets Angle Value as ``1/100 degree``"""
        ...
