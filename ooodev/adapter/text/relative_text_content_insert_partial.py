from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.text import XRelativeTextContentInsert

from ooodev.exceptions import ex as mEx
from ooodev.utils import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.text import XTextContent
    from ooodev.utils.type_var import UnoInterface


class RelativeTextContentInsertPartial:
    """
    Partial class for XTextRange.
    """

    # pylint: disable=unused-argument

    def __init__(
        self, component: XRelativeTextContentInsert, interface: UnoInterface | None = XRelativeTextContentInsert
    ) -> None:
        """
        Constructor

        Args:
            component (XRelativeTextContentInsert): UNO Component that implements ``com.sun.star.text.XRelativeTextContentInsert`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XTextRange``.
        """

        self.__interface = interface
        self.__validate(component)
        self.__component = component

    def __validate(self, component: Any) -> None:
        """
        Validates the component.

        Args:
            component (Any): The component to be validated.
        """
        if self.__interface is None:
            return
        if not mLo.Lo.is_uno_interfaces(component, self.__interface):
            raise mEx.MissingInterfaceError(self.__interface)

    # region XRelativeTextContentInsert
    def insert_text_content_after(self, new_content: XTextContent, predecessor: XTextContent) -> None:
        """
        Inserts text the new text content after the predecessor argument.

        This is helpful to insert paragraphs after text tables especially in headers, footers or text frames.

        Args:
            new_content (XTextContent): The new text content.
            predecessor (XTextContent): The predecessor text content.

        Raises:
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
        """
        self.__component.insertTextContentAfter(new_content, predecessor)

    def insert_text_content_before(self, new_content: XTextContent, successor: XTextContent) -> None:
        """
        inserts text the new text content before of the successor argument.

        This is helpful to insert paragraphs before of text tables.

        Args:
            new_content (XTextContent): The new text content.
            successor (XTextContent): The successor text content.

        Raises:
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
        """
        self.__component.insertTextContentBefore(new_content, successor)

    # endregion XRelativeTextContentInsert
