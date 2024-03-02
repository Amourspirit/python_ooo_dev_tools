from __future__ import annotations
from typing import cast, TYPE_CHECKING, Tuple
from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.text.text_section_partial import TextSectionPartial


if TYPE_CHECKING:
    from com.sun.star.text import TextSection  # service
    from com.sun.star.text import XTextSection


class TextSectionComp(ComponentBase, TextSectionPartial):
    """
    Class for managing TextSection Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XTextSection) -> None:
        """
        Constructor

        Args:
            component (TextSection): UNO TextRange Component that supports ``com.sun.star.text.TextSection`` service.
        """

        ComponentBase.__init__(self, component)
        TextSectionPartial.__init__(self, component)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        # validated by TextSectionPartial
        return ()  # ("com.sun.star.text.TextSection",)

    def get_child_sections(self) -> Tuple[TextSectionComp, ...]:
        """
        Gets all text sections that are children of this text section (recursive).
        """
        result = []
        sections = self.component.getChildSections()
        if sections is not None:
            result.extend(TextSectionComp(section) for section in sections)
        return tuple(result)

    # endregion Overrides

    # region Properties
    @property
    def component(self) -> TextSection:
        """TextSection Component"""
        # pylint: disable=no-member
        return cast("TextSection", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
