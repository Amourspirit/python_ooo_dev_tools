# region imports
from __future__ import annotations
from collections import defaultdict
from typing import Any, cast, Dict, Sequence, TYPE_CHECKING
import uno  # pylint: disable=unused-import

from com.sun.star.awt.tree import XMutableTreeDataModel

from ooodev.adapter.view.selection_change_events import SelectionChangeEvents
from ooodev.adapter.awt.tree.tree_edit_events import TreeEditEvents
from ooodev.adapter.awt.tree.tree_expansion_events import TreeExpansionEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils import lo as mLo
from ooodev.dialog.search.tree_search import (
    SearchTree,
    RuleDataCompare,
    RuleDataInsensitive,
    RuleTextSensitive,
    RuleTextInsensitive,
)
from .ctl_base import DialogControlBase


if TYPE_CHECKING:
    from com.sun.star.awt.tree import TreeControl  # service
    from com.sun.star.awt.tree import TreeControlModel  # service
    from com.sun.star.awt.tree import XMutableTreeNode
    from com.sun.star.awt.tree import XTreeNode
# endregion imports


class CtlTree(DialogControlBase, SelectionChangeEvents, TreeEditEvents, TreeExpansionEvents):
    """Class for Tree Control"""

    # The API docs does not show it but the Tree Control does support the standard UNO events in CtlListenerBase.

    # pylint: disable=unused-argument

    # region init
    def __init__(self, ctl: TreeControl) -> None:
        """
        Constructor

        Args:
            ctl (TreeControl): Tree Control
        """
        # generally speaking EventArgs.event_data will contain the Event object for the UNO event raised.
        DialogControlBase.__init__(self, ctl)
        generic_args = self._get_generic_args()
        # EventArgs.event_data will contain the ActionEvent
        SelectionChangeEvents.__init__(
            self, trigger_args=generic_args, cb=self._on_selection_change_events_listener_add_remove
        )
        TreeEditEvents.__init__(self, trigger_args=generic_args, cb=self._on_tree_edit_events_listener_add_remove)
        TreeExpansionEvents.__init__(
            self, trigger_args=generic_args, cb=self._on_tree_expansion_events_listener_add_remove
        )

    # endregion init

    # region Lazy Listeners
    def _on_selection_change_events_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        key = cast(str, event.source)
        if self._has_listener(key):
            return
        self.view.addSelectionChangeListener(self.events_listener_selection_change)
        self._add_listener(key)

    def _on_tree_edit_events_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        key = cast(str, event.source)
        if self._has_listener(key):
            return
        self.view.addTreeEditListener(self.events_listener_tree_edit)
        self._add_listener(key)

    def _on_tree_expansion_events_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        key = cast(str, event.source)
        if self._has_listener(key):
            return
        self.view.addTreeExpansionListener(self.events_listener_tree_expansion)
        self._add_listener(key)

    # endregion Lazy Listeners

    # region Overrides
    def get_view_ctl(self) -> TreeControl:
        return cast("TreeControl", super().get_view_ctl())

    def get_uno_srv_name(self) -> str:
        """Returns ``com.sun.star.awt.TreeControl``"""
        return "com.sun.star.awt.TreeControl"

    def get_model(self) -> TreeControlModel:
        """Gets the Model for the control"""
        # Tree Control does have getModel() method even thought it is not documented
        return cast("TreeControlModel", self.get_view_ctl().getModel())  # type: ignore

    # endregion Overrides

    # region Tree Nodes
    def create_root(self, display_value: str, data_value: Any = None) -> XMutableTreeNode:
        """
        Creates a root node for the tree

        Args:
            display_value (str): Display value for the root node.
            data_value (Any, optional): Specifies any value associated with the root node. Defaults to None.

        Returns:
            XMutableTreeNode: Returns a new root node of the tree control.
        """
        if not self.model.DataModel:
            raise ValueError("DataModel is not set")
        dm = mLo.Lo.qi(XMutableTreeDataModel, self.model.DataModel, True)
        root = dm.createNode(display_value, True)
        if data_value is not None:
            root.DataValue = data_value
        dm.setRoot(root)

        # To be visible, a root must have contained at least 1 child. Create a fictive one and erase it.
        # This behavior does not seem related to the RootDisplayed property
        root.appendChild(dm.createNode("Fictive", False))
        root.removeChildByIndex(0)
        return root

    def add_sub_node(
        self, parent_node: XMutableTreeNode, display_value: str = "", data_value: Any = None
    ) -> XMutableTreeNode:
        """
        Adds a sub node to the parent node

        Args:
            parent_node (XMutableTreeNode): Parent node
            display_value (str, optional): display value for the Node.
            data_value (Any, optional): Specifies any value associated with the node. Defaults to None.

        Raises:
            ValueError: _description_

        Returns:
            XMutableTreeNode: _description_
        """
        if not self.model.DataModel:
            raise ValueError("DataModel is not set")
        dm = mLo.Lo.qi(XMutableTreeDataModel, self.model.DataModel, True)
        node = dm.createNode(display_value, True)
        if data_value is not None:
            node.DataValue = data_value

        parent_node.appendChild(node)
        return node

    def add_sub_tree(
        self, flat_tree: Sequence[Sequence[str]], parent_node: XMutableTreeNode | None = None, width_data: bool = False
    ) -> None:
        """
        Adds a sub tree to the parent node

        Args:
            parent_node (XMutableTreeNode): Parent node.
            flat_tree (Sequence[Sequence[str]]): FlatTree: a 2D sequence of strings, sorted on the columns containing the DisplayValues
            width_data (bool, optional): _description_. Defaults to False.

        Raises:
            ValueError: _description_
        """
        if not flat_tree:
            return
        tree_data = self.convert_to_tree(flat_tree)
        self.add_nodes_from_tree_data(tree_data, parent_node)

    def add_nodes_from_tree_data(self, tree_data: Dict[str, str], parent_node: XMutableTreeNode | None = None) -> None:
        """
        Adds nodes to the control from a tree data structure.

        Args:
            tree_data (Dict[str, str]): A tree data structure.
            parent_node (XMutableTreeNode, optional): The parent node to add the nodes to. If None, adds nodes to the root of the control. Defaults to None.
        """
        # if parent_node is None and add_root_nodes is False:
        #     parent_node = self.create_root("Root")

        for key, value in tree_data.items():
            if parent_node is None:
                node = self.create_root(key)
            else:
                node = self.add_sub_node(parent_node, key)
            if isinstance(value, dict):
                self.add_nodes_from_tree_data(value, node)
            else:
                self.add_sub_node(node, value)

    def convert_to_tree(self, flat_tree: Sequence[Sequence[str]]) -> Dict[str, str]:
        """
        Converts a flat tree to a tree

        Args:
            flat_tree (Sequence[Sequence[str]]): FlatTree: a 2D sequence of strings, sorted on the columns containing the DisplayValues

        Returns:
            Dict[str, str]: A tree
        """

        def tree():
            return defaultdict(tree)

        def add(t, path):
            for node in path:
                t = t[node]

        root = tree()
        for path in flat_tree:
            add(root, path)
        return root

    def find_node(
        self, node: XTreeNode, value: str, case_sensitive: bool = False, search_data_value: bool = True
    ) -> XTreeNode | None:
        """
        Perform a search on a tree from a given node, looking for a node with a specific value.

        Args:
            node (XTreeNode): Node to start search from.
            value (str): Value to search for.
            case_sensitive (bool, optional): Specifies if the search is case sensitive. Defaults to ``False``.
            search_data_value (bool, optional): Specifies if ``DataValue`` of nodes are to be include in search. Defaults to ``True``.

        Returns:
            XTreeNode | None: Tree node if found; Otherwise, None.

        Note:
            :py:class:`~ooodev.dialog.search.tree_search.SearchTree` is a much more powerful search tool.
            It can be used to search for other types of match such as regular expressions.

            Custom rules can be created if the exiting rules do no cover you search needs.

        See Also:
            :ref:`ns_dialog_search_tree_search_`
        """
        # Check the current node
        search = SearchTree(match_value=value, match_all=False)
        if case_sensitive:
            search.register_rule(RuleTextSensitive())
            if search_data_value:
                search.register_rule(RuleDataCompare("="))
        else:
            search.register_rule(RuleTextInsensitive())
            if search_data_value:
                search.register_rule(RuleDataInsensitive())

        return search.find_node(node)

    # endregion Tree Nodes

    # region Properties
    @property
    def view(self) -> TreeControl:
        return self.get_view_ctl()

    @property
    def model(self) -> TreeControlModel:
        return self.get_model()

    # endregion Properties
