from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING

try:
    # python 3.12+
    from typing import override  # noqa # type: ignore
except ImportError:
    from typing_extensions import override  # noqa # type: ignore

from ooodev.adapter.component_base import ComponentBase

from ooodev.adapter.container.enumeration_access_partial import EnumerationAccessPartial
from ooodev.adapter.text.text_partial import TextPartial

if TYPE_CHECKING:
    from com.sun.star.text import Text


class TextComp(ComponentBase, EnumerationAccessPartial["Paragraph"], TextPartial):
    """
    Class for managing Text Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: Any) -> None:
        """
        Constructor

        Args:
            component (Any): UNO Component that supports ``com.sun.star.text.Text`` service.
        """

        ComponentBase.__init__(self, component)
        EnumerationAccessPartial.__init__(self, component, interface=None)
        TextPartial.__init__(self, component, interface=None)

    # region Overrides
    @override
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.text.Text",)

    # endregion Overrides

    # region XEnumerationAccess

    # endregion XEnumerationAccess

    # region Properties
    @property
    @override
    def component(self) -> Text:
        """Text Component"""
        # pylint: disable=no-member
        return cast("Text", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
