from __future__ import annotations
from typing import cast, TYPE_CHECKING
import uno
from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.container.element_access_partial import ElementAccessPartial
from ooodev.adapter.style.paragraph_properties_partial import ParagraphPropertiesPartial
from ooodev.adapter.style.character_properties_partial import CharacterPropertiesPartial


if TYPE_CHECKING:
    from com.sun.star.text import Paragraph


class ParagraphComp(ComponentBase, ParagraphPropertiesPartial, CharacterPropertiesPartial, ElementAccessPartial):
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

        ComponentBase.__init__(self, component=component)
        ParagraphPropertiesPartial.__init__(self, component=component)
        CharacterPropertiesPartial.__init__(self, component=component)
        ElementAccessPartial.__init__(self, component=component, interface=None)

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
