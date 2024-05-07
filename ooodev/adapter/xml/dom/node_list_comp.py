from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno
from com.sun.star.xml.dom import XNodeList
from ooodev.adapter.component_prop import ComponentProp
from ooodev.adapter.xml.dom.node_list_partial import NodeListPartial
from ooodev.utils import gen_util as mGenUtil

if TYPE_CHECKING:
    from com.sun.star.xml.dom import XNode


class NodeListComp(ComponentProp, NodeListPartial):
    """
    Class for managing NodeListPartial Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XNodeList) -> None:
        """
        Constructor

        Args:
            component (NodeListPartial): UNO Component that supports ``com.sun.star.util.NodeListPartial`` service.
        """
        # pylint: disable=no-member
        ComponentProp.__init__(self, component)
        NodeListPartial.__init__(self, component=component, interface=None)
        self.__current_index = -1

    def _is_next_index_element_valid(self, element: Any) -> bool:
        """
        Gets if the next element is valid.
        This method is called when iterating over the elements of this class.

        Args:
            element (Any): Element

        Returns:
            bool: True in this class but can be overridden in child classes.
        """
        return True

    def __iter__(self):
        """
        Iterates over the nodes.

        Yields:
            XNode: Node
        """
        self.__current_index = 0
        while self.__current_index < len(self):
            yield self.component.item(self.__current_index)
            self.__current_index += 1

    def __next__(self):
        """
        Gets the next node.

        Returns:
            XNode: The next node.
        """
        if self.__current_index > len(self):
            raise StopIteration
        else:
            self.__current_index += 1
            return self.component.item(self.__current_index - 1)

    def __reversed__(self):
        """
        Iterates over the nodes in reverse.

        Yields:
            XNode: Node
        """
        self.__current_index = len(self) - 1
        while self.__current_index >= 0:
            yield self.component.item(self.__current_index)
            self.__current_index -= 1

    def __len__(self) -> int:
        """
        Gets the number of nodes in the list.

        Returns:
            int: Number of nodes in the list.
        """
        return self.component.getLength()

    def __getitem__(self, idx: int) -> XNode:
        """
        Gets the node at the specified index.

        Args:
            key (idx, int): The index of the node. When getting by index can be a negative value to get from the end.

        Returns:
            XNode: The node at the specified index.
        """
        count = len(self)
        index = mGenUtil.Util.get_index(idx, count, False)
        return self.component.item(index)

    # region XNodeList Overrides
    def item(self, idx: int) -> XNode:
        """
        Returns a node specified by index in the collection.

        Args:
            idx (int): Index of node. When getting by index can be a negative value to get from the end.

        Returns:
            XNode: The node at the specified index.
        """
        return self[idx]

    # endregion XNodeList Overrides

    # region Properties
    @property
    def component(self) -> XNodeList:
        """XNodeList Component"""
        # pylint: disable=no-member
        return cast("XNodeList", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
