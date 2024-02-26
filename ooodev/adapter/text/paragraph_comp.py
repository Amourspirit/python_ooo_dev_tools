from __future__ import annotations
from typing import cast, TYPE_CHECKING
import uno
from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.container.element_access_partial import ElementAccessPartial


if TYPE_CHECKING:
    from com.sun.star.text import Paragraph


class ParagraphComp(ComponentBase, ElementAccessPartial):
    """
    Class for managing Paragraph Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: Paragraph) -> None:
        """
        Constructor

        Args:
            component (Paragraph): UNO TextContent Component that supports ``com.sun.star.text.Paragraph`` service.
        """

        ComponentBase.__init__(self, component)
        ElementAccessPartial.__init__(self, component)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.text.Paragraph",)

    # endregion Overrides

    # region XEnumerationAccess

    # endregion XEnumerationAccess

    # region Properties
    @property
    def component(self) -> Paragraph:
        """Paragraph Component"""
        # pylint: disable=no-member
        return cast("Paragraph", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
