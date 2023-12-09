from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.text import XTextContent

from ooodev.exceptions import ex as mEx
from ooodev.utils import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.text import XTextRange
    from ooodev.utils.type_var import UnoInterface


class TextContentPartial:
    """
    Partial class for XTextContent.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XTextContent, interface: UnoInterface | None = XTextContent) -> None:
        """
        Constructor

        Args:
            component (XTextContent): UNO Component that implements ``com.sun.star.text.XTextContent`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XTextContent``.
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

    # region XTextContent
    def attach(self, text_range: XTextRange) -> None:
        """Attaches a text range to this text content."""
        self.__component.attach(text_range)

    def get_anchor(self) -> XTextRange:
        """Returns the anchor of this text content."""
        return self.__component.getAnchor()

    # endregion XTextContent
