from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.util import XStringEscape

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class StringEscapePartial:
    """
    Partial class for XStringEscape.
    """

    def __init__(self, component: XStringEscape, interface: UnoInterface | None = XStringEscape) -> None:
        """
        Constructor

        Args:
            component (XStringEscape): UNO Component that implements ``com.sun.star.util.XStringEscape.XStringEscape`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XStringEscape``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XStringEscape

    def escape_string(self, string: str) -> str:
        """
        Encodes an arbitrary string into an escaped form compatible with some naming rules.

        Raises:
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
        """
        return self.__component.escapeString(string)

    def unescapeString(self, escaped_string: str) -> str:
        """
        decodes an escaped string into the original form.

        Raises:
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
        """
        return self.__component.unescapeString(escaped_string)

    # endregion XStringEscape


def get_builder(component: Any) -> Any:
    """
    Get the builder for the component.

    Args:
        component (Any): The component.

    Returns:
        DefaultBuilder: Builder instance.
    """
    # pylint: disable=import-outside-toplevel
    from ooodev.utils.builder.default_builder import DefaultBuilder

    builder = DefaultBuilder(component)
    builder.add_import(
        name="ooodev.adapter.util.string_escape_partial.StringEscapePartial",
        uno_name="com.sun.star.util.XStringEscape",
        optional=False,
        init_kind=2,
    )
    return builder
