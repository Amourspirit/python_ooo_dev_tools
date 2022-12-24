from __future__ import annotations
from typing import Any, Tuple, TYPE_CHECKING
import uno
from com.sun.star.beans import XPropertySet
from ..utils import props as mProps

try:
    from typing import Protocol
except ImportError:
    from typing_extensions import Protocol
if TYPE_CHECKING:
    from com.sun.star.beans import PropertyValue


class StyleObj(Protocol):
    """
    Protolcol Class for Styles
    """

    def apply_style(self, obj: object) -> None:
        """Applies style"""
        ...

    def get_props(self) -> Tuple[PropertyValue, ...]:
        """Gets Properties"""
        ...
