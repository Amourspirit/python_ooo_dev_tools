from __future__ import annotations
from typing import cast, TYPE_CHECKING
import uno

from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.style.paragraph_properties_partial import ParagraphPropertiesPartial

if TYPE_CHECKING:
    from com.sun.star.style import ParagraphProperties  # service


class ParagraphPropertiesComp(ComponentBase, ParagraphPropertiesPartial):
    """
    Class for managing table ParagraphProperties Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: ParagraphProperties) -> None:
        """
        Constructor

        Args:
            component (ParagraphProperties): UNO ParagraphProperties Component.
        """
        ComponentBase.__init__(self, component)
        ParagraphPropertiesPartial.__init__(self, component=component)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.drawing.TextProperties", "com.sun.star.style.ParagraphProperties")

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> ParagraphProperties:
        """ParagraphProperties Component"""
        # pylint: disable=no-member
        return cast("ParagraphProperties", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
