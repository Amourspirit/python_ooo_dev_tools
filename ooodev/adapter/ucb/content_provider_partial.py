from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno
from com.sun.star.ucb import XContentProvider

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.ucb import XContentIdentifier
    from com.sun.star.ucb import XContent
    from ooodev.utils.type_var import UnoInterface


class ContentProviderPartial:
    """
    Partial Class XContentProvider.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XContentProvider, interface: UnoInterface | None = XContentProvider) -> None:
        """
        Constructor

        Args:
            component (XContentProvider): UNO Component that implements ``com.sun.star.ucb.XContentProvider`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XContentProvider``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XContentProvider
    def compare_content_ids(self, id1: XContentIdentifier, id2: XContentIdentifier) -> int:
        """
        compares two XContentIdentifiers.
        """
        return self.__component.compareContentIds(id1, id2)

    def query_content(self, identifier: XContentIdentifier) -> XContent:
        """
        creates a new XContent instance, if the given XContentIdentifier matches a content provided by the implementation of this interface.

        Raises:
            com.sun.star.ucb.IllegalIdentifierException: ``IllegalIdentifierException``
        """
        return self.__component.queryContent(identifier)

    # endregion XContentProvider
