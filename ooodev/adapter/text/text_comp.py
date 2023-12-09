from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno

from ooodev.adapter.component_base import ComponentBase

from ooodev.adapter.container.enumeration_access_partial import EnumerationAccessPartial

if TYPE_CHECKING:
    from com.sun.star.container import XEnumerationAccess
    from com.sun.star.text import Text


class TextComp(ComponentBase, EnumerationAccessPartial):
    """
    Class for managing Text Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XEnumerationAccess) -> None:
        """
        Constructor

        Args:
            component (Text): UNO TextContent Component that supports ``com.sun.star.text.Text`` service.
        """

        ComponentBase.__init__(self, component)
        EnumerationAccessPartial.__init__(self, component)

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
