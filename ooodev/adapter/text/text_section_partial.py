from __future__ import annotations
from typing import TYPE_CHECKING, Tuple
import uno

from com.sun.star.text import XTextSection

from ooodev.adapter.text.text_content_partial import TextContentPartial

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class TextSectionPartial(TextContentPartial):
    """
    Partial class for XTextSection.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XTextSection, interface: UnoInterface | None = XTextSection) -> None:
        """
        Constructor

        Args:
            component (XTextSection): UNO Component that implements ``com.sun.star.text.XTextSection`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XTextSection``.
        """
        TextContentPartial.__init__(self, component, interface=interface)
        self.__component = component

    # region XTextSection
    def get_child_sections(self) -> Tuple[XTextSection, ...]:
        """
        Gets all text sections that are children of this text section (recursive).
        """
        result = self.__component.getChildSections()
        return () if result is None else result

    def get_parent_section(self) -> XTextSection | None:
        """
        If this instance is a child section, then this method returns the parent text section.
        """
        return self.__component.getParentSection()

    # endregion XTextSection
