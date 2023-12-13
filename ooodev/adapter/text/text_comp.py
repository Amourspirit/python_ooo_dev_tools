from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno

from ooodev.adapter.component_base import ComponentBase

from ooodev.adapter.container.enumeration_access_partial import EnumerationAccessPartial
from .text_partial import TextPartial

if TYPE_CHECKING:
    from com.sun.star.container import XEnumerationAccess
    from com.sun.star.text import Text


class TextComp(ComponentBase, EnumerationAccessPartial, TextPartial):
    """
    Class for managing Text Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: Any) -> None:
        """
        Constructor

        Args:
            component (Any): UNO TextContent Component that supports ``com.sun.star.text.Text`` service.
        """

        ComponentBase.__init__(self, component)
        EnumerationAccessPartial.__init__(self, component, interface=None)
        TextPartial.__init__(self, component, interface=None)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.text.Text",)

    # endregion Overrides

    # region XEnumerationAccess

    # endregion XEnumerationAccess

    # region Properties
    @property
    def component(self) -> Text:
        """Sheet Cell Cursor Component"""
        return cast("Text", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
