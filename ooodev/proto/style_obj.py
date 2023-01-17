from __future__ import annotations
from typing import Tuple, TYPE_CHECKING
from ..format.kind.format_kind import FormatKind as FormatKind

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

    def backup(self, obj: object) -> None:
        """
        Backs up Attriubes that are to be changed by apply.

        If used method should be called before apply.

        Args:
            obj (object): Object to backup properties from.

        Returns:
            None:
        """
        ...

    def restore(self, obj: object, clear: bool = False) -> None:
        """
        Restores ``obj`` properties from backed up setting if any exist.

        Restore can only be effective if ``backup()`` has be run before calling this method.

        Args:
            obj (object): Object to restore properties on.
            clear (bool): Determines if backup is cleared after resore. Default ``False``

        Returns:
            None:

        See Also:
            :py:meth:`~.style_base.StyleBase.backup`
        """
        ...

    @property
    def prop_has_attribs(self) -> bool:
        """Gets If instantance has any attributes set."""
        ...

    @property
    def prop_has_backup(self) -> bool:
        """Gets If instantance has backup data set."""
        ...

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        ...
