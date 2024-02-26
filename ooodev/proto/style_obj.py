from __future__ import annotations
from typing import Tuple, TYPE_CHECKING, Any
from ooodev.format.inner.kind.format_kind import FormatKind as FormatKind

if TYPE_CHECKING:
    from com.sun.star.beans import PropertyValue

    from typing_extensions import Protocol
else:
    Protocol = object


class StyleT(Protocol):
    """
    Protocol Class for Styles
    """

    def apply(self, obj: Any) -> None:
        """Applies style"""
        ...

    # @overload
    # def apply(self, obj: Any, **kwargs: Any) -> None:
    #     """Applies style"""
    #     ...

    def get_props(self) -> Tuple[PropertyValue, ...]:
        """Gets Properties"""
        ...

    def get_attrs(self) -> Tuple[str, ...]:
        """
        Gets the attributes that are slated for change in the current instance

        Returns:
            Tuple(str, ...): Tuple of Attributes
        """
        ...

    def backup(self, obj: Any) -> None:
        """
        Backs up Attributes that are to be changed by apply.

        If used method should be called before apply.

        Args:
            obj (object): Object to backup properties from.

        Returns:
            None:
        """
        ...

    def restore(self, obj: Any, clear: bool = False) -> None:
        """
        Restores ``obj`` properties from backed up setting if any exist.

        Restore can only be effective if ``backup()`` has be run before calling this method.

        Args:
            obj (object): Object to restore properties on.
            clear (bool): Determines if backup is cleared after restore. Default ``False``

        Returns:
            None:

        See Also:
            :py:meth:`~.style_base.StyleBase.backup`
        """
        ...

    def support_service(self, *service: str) -> bool:
        """
        Gets if service is supported.

        Args:
            service: expandable list of service names of UNO services such as ``com.sun.star.text.TextFrame``.

        Returns:
            bool: ``True`` if service is supported; Otherwise, ``False``.
        """
        ...

    def set_update_obj(self, obj: Any) -> None:
        """
        Sets the update object for the style instance.

        Args:
            obj (Any): Object used to apply style to when update is called.

        Returns:
            None:
        """
        ...

    # don't know why but pyright complains about this if properties are specific types.
    # May be a bug in pyright. May only be for child classes.

    @property
    def prop_has_backup(self) -> Any:
        """Gets If instance has backup data set."""
        ...

    @property
    def prop_has_attribs(self) -> Any:
        """Gets If instance has any attributes set."""
        ...

    @property
    def prop_format_kind(self) -> Any:
        """Gets the kind of style"""
        ...


# StyleOby was renamed to StyleT in version 0.13.8
StyleObj = StyleT
