from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING

from ooodev.events.args.generic_args import GenericArgs
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils import gen_util as gUtil
from ooodev.adapter.tree.tree_data_model_listener import TreeDataModelListener

if TYPE_CHECKING:
    from com.sun.star.awt.tree import XTreeDataModel
    from ooodev.utils.type_var import EventArgsCallbackT, ListenerEventCallbackT


class TreeDataModelEvents:
    """
    Class for managing Tree Data Model Events.
    """

    def __init__(
        self,
        trigger_args: GenericArgs | None = None,
        cb: ListenerEventCallbackT | None = None,
        listener: TreeDataModelListener | None = None,
    ) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
                This only applies if the listener is not passed.
            cb (ListenerEventCallbackT | None, optional): Callback that is invoked when an event is added or removed.
            listener (TreeDataModelListener | None, optional): Listener that is used to manage events.
        """
        self.__callback = cb
        if listener:
            self.__listener = listener
        else:
            self.__listener = TreeDataModelListener(trigger_args=trigger_args)
        self.__name = gUtil.Util.generate_random_string(10)

    # region Manage Events
    def add_event_tree_nodes_changed(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked after a node (or a set of siblings) has changed in some way.

        The node(s) have not changed locations in the tree or altered their children arrays,
        but other attributes have changed and may affect presentation.

        Example: the name of a file has changed, but it is in the same location in the file system.

        To indicate the root has changed, TreeDataModelEvent.Nodes will contain the root
        node and ``TreeDataModelEvent.ParentNode`` will be empty.

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.awt.tree.TreeDataModelEvent`` struct.
        """
        # sourcery skip: class-extract-method
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="treeNodesChanged")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("treeNodesChanged", cb)

    def add_event_tree_nodes_inserted(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked after nodes have been inserted into the tree.

        Use ``TreeDataModelEvent.ParentNode`` to get the parent of the new node(s).
        ``TreeDataModelEvent.Nodes`` contains the new node(s).

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.awt.tree.TreeDataModelEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="treeNodesInserted")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("treeNodesInserted", cb)

    def add_event_tree_nodes_removed(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked after nodes have been removed from the tree.

        Note that if a subtree is removed from the tree, this method may only be invoked once for the root of the
        removed subtree, not once for each individual set of siblings removed.

        Use ``TreeDataModelEvent.ParentNode`` to get the former parent of the deleted node(s).
        ``TreeDataModelEvent.Nodes`` contains the removed node(s).

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.awt.tree.TreeDataModelEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="treeNodesRemoved")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("treeNodesRemoved", cb)

    def add_event_tree_structure_changed(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked after the tree has drastically changed structure from a given node down.

        Use ``TreeDataModelEvent.ParentNode`` to get the node which structure has changed.
        ``TreeDataModelEvent.Nodes`` is empty.

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.awt.tree.TreeDataModelEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="treeStructureChanged")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("treeStructureChanged", cb)

    def add_event_tree_data_model_events_disposing(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when the broadcaster is about to be disposed.

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.lang.EventObject`` struct.
        """

        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="disposing")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("disposing", cb)

    def remove_event_tree_nodes_changed(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="treeNodesChanged", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("treeNodesChanged", cb)

    def remove_event_tree_nodes_inserted(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="treeNodesInserted", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("treeNodesInserted", cb)

    def remove_event_tree_nodes_removed(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="treeNodesRemoved", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("treeNodesRemoved", cb)

    def remove_event_tree_structure_changed(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="treeStructureChanged", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("treeStructureChanged", cb)

    def remove_event_tree_data_model_events_disposing(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="disposing", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("disposing", cb)

    @property
    def events_listener_tree_data_model(self) -> TreeDataModelListener:
        """
        Returns listener
        """
        return self.__listener

    # endregion Manage Events


def on_lazy_cb(source: Any, event: ListenerEventArgs) -> None:
    """
    Callback that is invoked when an event is added or removed.

    This method is generally used to add the listener to the component in a lazy manner.
    This means this callback will only be called once in the lifetime of the component.

    Args:
        source (Any): Expected to be an instance of TreeDataModelEvents that is a partial class of a component based class.
        event (ListenerEventArgs): Event arguments.

    Returns:
        None:

    Warning:
        This method is intended for internal use only.
    """
    # will only ever fire once
    if not isinstance(source, TreeDataModelEvents):
        return
    if not hasattr(source, "component"):
        return

    comp = cast("XTreeDataModel", source.component)  # type: ignore
    comp.addTreeDataModelListener(source.events_listener_tree_data_model)
    event.remove_callback = True
