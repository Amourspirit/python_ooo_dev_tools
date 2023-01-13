from __future__ import annotations
from typing import Tuple, TYPE_CHECKING
from ..format.kind.style_kind import StyleKind as StyleKind

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

    def apply(self, obj: object, **kwargs) -> None:
        """Applies style"""
        ...

    def get_props(self) -> Tuple[PropertyValue, ...]:
        """Gets Properties"""
        ...

    def get_attrs(self) -> Tuple[str, ...]:
        """
        Gets the attributes that are slated for change in the current instance

        Returns:
            Tuple(str, ...): Tuple of attribures
        """
        ...

    @property
    def prop_has_attribs(self) -> bool:
        """Gets If instantance has any attributes set."""
        ...

    @property
    def prop_style_kind(self) -> StyleKind:
        """Gets the kind of style"""
        ...
