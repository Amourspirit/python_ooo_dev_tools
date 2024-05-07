from __future__ import annotations
from typing import Any, TYPE_CHECKING

import uno
from com.sun.star.xml.dom import XNodeList

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.xml.dom import XNode
    from ooodev.utils.type_var import UnoInterface


class NodeListPartial:
    """
    Partial class for XNodeList.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XNodeList, interface: UnoInterface | None = XNodeList) -> None:
        """
        Constructor

        Args:
            component (XNodeList ): UNO Component that implements ``com.sun.star.xml.dom.XNodeList`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XNodeList``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XNodeList
    def get_length(self) -> int:
        """
        The number of nodes in the list.
        """
        return self.__component.getLength()

    def item(self, idx: int) -> XNode:
        """
        Returns a node specified by index in the collection.
        """
        return self.__component.item(idx)

    # endregion XNodeList
