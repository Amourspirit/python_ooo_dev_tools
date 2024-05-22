from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno
from com.sun.star.ucb import XContent

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.ucb import XContentEventListener
    from com.sun.star.ucb import XContentIdentifier
    from ooodev.utils.type_var import UnoInterface


class ContentPartial:
    """
    Partial Class XContent.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XContent, interface: UnoInterface | None = XContent) -> None:
        """
        Constructor

        Args:
            component (XContent): UNO Component that implements ``com.sun.star.ucb.XContent`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XContent``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XContent
    def add_content_event_listener(self, listener: XContentEventListener) -> None:
        """
        Adds a listener for content events.
        """
        self.__component.addContentEventListener(listener)

    def get_content_type(self) -> str:
        """
        Returns a type string, which is unique for that type of content (e.g. ``application/vnd.sun.star.hierarchy-folder``).
        """
        return self.__component.getContentType()

    def get_identifier(self) -> XContentIdentifier:
        """
        Returns the identifier of the content.
        """
        return self.__component.getIdentifier()

    def remove_content_event_listener(self, listener: XContentEventListener) -> None:
        """
        Removes a listener for content events.
        """
        self.__component.removeContentEventListener(listener)

    # endregion XContent
