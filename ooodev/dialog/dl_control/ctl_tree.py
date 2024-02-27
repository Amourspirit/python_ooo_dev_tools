# region imports
from __future__ import annotations
from collections import defaultdict
import contextlib
from typing import Any, List, cast, Sequence, TYPE_CHECKING
import uno  # pylint: disable=unused-import

from com.sun.star.awt.tree import XMutableTreeDataModel
from ooo.dyn.view.selection_type import SelectionType  # enum

from ooodev.adapter.awt.tree.tree_edit_events import TreeEditEvents
from ooodev.adapter.awt.tree.tree_expansion_events import TreeExpansionEvents
from ooodev.adapter.tree.tree_data_model_comp import TreeDataModelComp
from ooodev.adapter.view.selection_change_events import SelectionChangeEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.loader import lo as mLo
from ooodev.utils.kind.dialog_control_kind import DialogControlKind
from ooodev.utils.kind.dialog_control_named_kind import DialogControlNamedKind
from ooodev.adapter.awt.tree.tree_control_model_partial import TreeControlModelPartial
from ooodev.dialog.search.tree_search.search_tree import SearchTree
from ooodev.dialog.search.tree_search.rule_data_compare import RuleDataCompare
from ooodev.dialog.search.tree_search.rule_data_insensitive import RuleDataInsensitive
from ooodev.dialog.search.tree_search.rule_text_sensitive import RuleTextSensitive
from ooodev.dialog.search.tree_search.rule_text_insensitive import RuleTextInsensitive
from ooodev.dialog.dl_control.ctl_base import DialogControlBase


if TYPE_CHECKING:
    from com.sun.star.awt.tree import MutableTreeNode  # service
    from com.sun.star.awt.tree import TreeControl  # service
    from com.sun.star.awt.tree import TreeControlModel  # service
    from com.sun.star.awt.tree import XMutableTreeNode
    from com.sun.star.awt.tree import XTreeNode
# endregion imports


class CtlTree(DialogControlBase, TreeControlModelPartial, SelectionChangeEvents, TreeEditEvents, TreeExpansionEvents):
    """Class for Tree Control"""

    DATA_VALUE_KEY = "___data_value___"

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
        TreeControlModelPartial.__init__(self)
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
        # will only ever fire once
        self.view.addSelectionChangeListener(self.events_listener_selection_change)
        event.remove_callback = True

    def _on_tree_edit_events_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.view.addTreeEditListener(self.events_listener_tree_edit)
        event.remove_callback = True

    def _on_tree_expansion_events_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.view.addTreeExpansionListener(self.events_listener_tree_expansion)
        event.remove_callback = True

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

    def get_control_kind(self) -> DialogControlKind:
        """Gets the control kind. Returns ``DialogControlKind.TREE``"""
        return DialogControlKind.TREE

    def get_control_named_kind(self) -> DialogControlNamedKind:
        """Gets the control named kind. Returns ``DialogControlNamedKind.TREE``"""
        return DialogControlNamedKind.TREE

    # endregion Overrides

    # region Tree Nodes
    def create_root(self, display_value: str, data_value: Any = None) -> XMutableTreeNode:
        """
        Creates a root node for the tree

        Args:
            display_value (str): Display value for the root node.
            data_value (Any, optional): Specifies any value associated with the node.
                Must be a type understood by UNO, such as a string, int, float, a struct such as ``UnoDateTime``, etc.
                Defaults to None.

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
            data_value (Any, optional): Specifies any value associated with the node.
                Must be a type understood by UNO, such as a string, int, float, a struct such as ``UnoDateTime``, etc.
                Defaults to None.

        Returns:
            XMutableTreeNode: MutableTreeNode
        """
        if not self.model.DataModel:
            raise ValueError("DataModel is not set")
        dm = mLo.Lo.qi(XMutableTreeDataModel, self.model.DataModel, True)
        node = dm.createNode(display_value, True)
        if data_value is not None:
            node.DataValue = data_value

        parent_node.appendChild(node)
        return node

    def add_sub_tree(self, flat_tree: Sequence[Any], parent_node: XMutableTreeNode | None = None) -> None:
        """
        Adds a sub tree to the parent node

        Args:
            parent_node (XMutableTreeNode): Parent node.
            flat_tree (Sequence[Sequence[str]]): FlatTree: a 2D sequence of strings, sorted on the columns containing the DisplayValues
            width_data (bool, optional): _description_. Defaults to False.

        Note:
            The same data structure for ``tree_data`` can be used to add sub-nodes as shown in :py:meth:`~.CtlTree.convert_to_tree`.
        """
        if not flat_tree:
            return
        tree_data = self.convert_to_tree(flat_tree)
        self._add_nodes_from_tree_data(tree_data, parent_node)

    def _add_nodes_from_tree_data(self, tree_data: dict, parent_node: XMutableTreeNode | None = None) -> None:
        """
        Adds nodes to the control from a tree data structure.

        Args:
            tree_data (Dict[str, str]): A tree data structure.
            parent_node (XMutableTreeNode, optional): The parent node to add the nodes to.
                If None, adds nodes to the root of the control. Defaults to None.
        """
        # if parent_node is None and add_root_nodes is False:
        #     parent_node = self.create_root("Root")

        def get_data_value(val: dict) -> Any:
            return val.get(CtlTree.DATA_VALUE_KEY, None)

        for key, value in tree_data.items():
            if key == CtlTree.DATA_VALUE_KEY:
                continue
            is_val_dict = isinstance(value, dict)
            if parent_node is None:
                if is_val_dict:
                    node = self.create_root(key, get_data_value(value))
                else:
                    node = self.create_root(key)
            else:
                if is_val_dict:
                    node = self.add_sub_node(parent_node, key, get_data_value(value))
                else:
                    node = self.add_sub_node(parent_node, key)
            if is_val_dict:
                self._add_nodes_from_tree_data(value, node)
            else:
                self.add_sub_node(node, value)

    def convert_to_tree(self, flat_tree: Sequence[Any]) -> dict:
        """
        Converts a flat tree to a tree

        Args:
            flat_tree (Sequence[Sequence[str]]): FlatTree: a 2D sequence of strings, sorted on the columns containing the DisplayValues

        Returns:
            dict: A tree

        Notes:
            The flat tree can be a sequence of sequence of strings or a sequence of sequence of sequence.

            **Example sequence of sequence of strings**:

            This example uses  a List of List of strings. It would alo work with a tuple of tuple of strings.

            .. code-block:: python

                [
                    ["A1", "B1", "C1"],
                    ["A1", "B1", "C2"],
                    ["A1", "B2", "C3"],
                    ["A2", "B3", "C4"],
                    ["A2", "B3", "C5"],
                    ["A2", "B3", "C6"],
                    ["A2", "B4", "Razor"],
                ]

            The result will be as follows:

            .. cssclass:: screen_shot

                .. image:: https://user-images.githubusercontent.com/4193389/283976149-e4763e71-c345-47fc-81d0-2ce86b93a8ce.png
                    :alt: Tree Control
                    :align: center

        **Example sequence of sequence of sequence**:

        This example uses includes data values that are to be assigned to the nodes.

        Data values can be any type understood by UNO, such as a string, int, float, a struct such as ``UnoDateTime``, etc.
        List and tuple can be interchanged and still work.

        In this example ``A1`` will have a data value of ``1`` and ``B1`` will have a data value of ``now_date``.
        The first data value that is encountered will be assigned to the node's ``DataValue`` property.
        All other data values for that node will be ignored.

        .. code-block:: python

            now_date = DateUtil.date_to_uno_date_time(datetime.datetime.now())

            [
                [("A1", 1), ("B1", now_date), ("C1", None)],
                [("A1", "ignored"), ("B1", None), ("C2", "Data4")],
                [("A1",), ("B2", "Data5"), ("C3", "Data6")],
                [("A2", 33), ("B3", "Data8"), ("C4", "Data9")],
                [("A2", "Data7"), ("B3", "Data8"), ("C5", "Data10")],
                [("A2", "Data7"), ("B3", "Data8"), ("C6", "Data11")],
            ]

        The result will be as follows:

        .. cssclass:: screen_shot

            .. image:: https://user-images.githubusercontent.com/4193389/283976966-ba27195e-58b7-4b98-8b16-ba64a86076e6.png
                :alt: Tree Control
                :align: center

        The ``B1`` Node will look something like this:

        .. cssclass:: screen_shot

            .. image:: https://user-images.githubusercontent.com/4193389/283977539-f22517ab-8b3e-4d42-8eb5-2425a0e3b065.png
                :alt: Tree Control
                :align: center

        The input is rather flexible. The following would also work:

        Note that ``A2`` contains too many values. The first two will be used and the rest ignored.
        The ``A2`` node will have a text value of ``A2`` and a data value of ``33``.

        .. code-block:: python

            [
                [["A1", 1], ["B1", now_date], ["C1"]],
                [["A1"], ["B1"], ["C2"]],
                [["A1"], ["B2"], ["C3"]],
                [["A2", 33, 66, 127], ["B3", "Data8"], ["C4", "Data9"]],
                [["A2"], ["B3", "Data8"], ["C5", "Data10"]],
                [["A2"], ["B3", None], ["C6", "Data11"]],
            ]

        See Also:
            :py:meth:`~.CtlTree.add_nodes_from_tree_data`
        """

        def tree():
            return defaultdict(tree)

        def get_lst(seq: Any, str_vals: bool) -> List[Any]:
            if str_vals:
                lst = [seq]
            else:
                lst = list(seq)
            while len(lst) < 2:
                lst.append(None)
            return lst

        def add(t, path, str_vals: bool):
            for seq in path:
                seq_len = len(seq)
                if seq_len == 0:
                    continue
                seq_lst = get_lst(seq, str_vals)
                node, data = seq_lst[:2]
                t = t[node]
                if data is not None and not CtlTree.DATA_VALUE_KEY in t:
                    t[CtlTree.DATA_VALUE_KEY] = data

        root = tree()
        if not flat_tree:
            return {}
        # get first item add check if it is a list or a string
        first = flat_tree[0]
        # ["A1", "B1", "C1"] or [("A1", "Data1"), ("B1", "Data2"), ("C1", "Data3")]
        if not first:
            return {}
        first_item = first[0]
        if isinstance(first_item, str):
            is_str_values = True
        else:
            is_str_values = False
        for path in flat_tree:
            add(root, path, is_str_values)
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
            :ref:`ns_dialog_search_tree_search`
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
    def current_selection(self) -> MutableTreeNode | None:
        """Gets the current selected node"""
        with contextlib.suppress(Exception):
            model = self.model
            if model.SelectionType != SelectionType.NONE:
                sel = self.view.getSelection()
                if sel is None:
                    return None
                if model.SelectionType == SelectionType.SINGLE:
                    return cast("MutableTreeNode", sel)  # expected to be MutableTreeNode
                if isinstance(sel, tuple) and len(sel) > 0:
                    return cast("MutableTreeNode", sel[0])
        return None

    # region TreeControlModelPartial overrides

    @property
    def data_model(self) -> TreeDataModelComp | None:
        """Gets the data model for the tree"""
        with contextlib.suppress(Exception):
            if not self.model.DataModel:
                return None
            return TreeDataModelComp(self.model.DataModel)
        return None

    # endregion TreeControlModelPartial overrides
    @property
    def model(self) -> TreeControlModel:
        # pylint: disable=no-member
        return cast("TreeControlModel", super().model)

    @property
    def root_node(self) -> MutableTreeNode | None:
        """Gets the root node of the tree"""
        with contextlib.suppress(Exception):
            if not self.model.DataModel:
                return None
            dm = mLo.Lo.qi(XMutableTreeDataModel, self.model.DataModel, True)
            return cast("MutableTreeNode", dm.getRoot())
        return None

    @property
    def view(self) -> TreeControl:
        # pylint: disable=no-member
        return cast("TreeControl", super().view)

    # endregion Properties
