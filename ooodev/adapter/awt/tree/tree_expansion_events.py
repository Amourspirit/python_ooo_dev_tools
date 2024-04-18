from __future__ import annotations

from typing import Any, cast, TYPE_CHECKING
from ooodev.events.args.generic_args import GenericArgs
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils import gen_util as gUtil
from ooodev.utils.type_var import EventArgsCallbackT, ListenerEventCallbackT
from ooodev.adapter.awt.tree.tree_expansion_listener import TreeExpansionListener

if TYPE_CHECKING:
    from com.sun.star.awt.tree import XTreeControl


class TreeExpansionEvents:
    """
    Class for managing Tree Expansion Events.

    This class is usually inherited by control classes that implement ``com.sun.star.awt.tree.XTreeExpansionListener``.
    """

    def __init__(
        self,
        trigger_args: GenericArgs | None = None,
        cb: ListenerEventCallbackT | None = None,
        listener: TreeExpansionListener | None = None,
        subscriber: XTreeControl | None = None,
    ) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
                This only applies if the listener is not passed.
            cb (ListenerEventCallbackT | None, optional): Callback that is invoked when an event is added or removed.
            listener (TreeExpansionListener | None, optional): Listener that is used to manage events.
            subscriber (XTreeControl, optional): An UNO object that implements the ``XTreeControl`` interface.
                If passed in then this instance listener is automatically added to it.
        """
        self.__callback = cb
        if listener:
            self.__listener = listener
            if subscriber:
                subscriber.addTreeExpansionListener(self.__listener)
        else:
            self.__listener = TreeExpansionListener(trigger_args=trigger_args, subscriber=subscriber)
        self.__name = gUtil.Util.generate_random_string(10)

    # region Manage Events
    def add_event_request_child_nodes(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when a node with children on demand is about to be expanded.

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.awt.tree.TreeExpansionEvent`` struct.
        """
        # sourcery skip: class-extract-method
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="requestChildNodes")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("requestChildNodes", cb)

    def add_event_tree_collapsed(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked whenever a node in the tree has been successfully collapsed.

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.awt.tree.TreeExpansionEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="treeCollapsed")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("treeCollapsed", cb)

    def add_event_tree_collapsing(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked whenever a node in the tree is about to be collapsed.

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.awt.tree.TreeExpansionEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="treeCollapsing")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("treeCollapsing", cb)

    def add_event_tree_expanded(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked whenever a node in the tree has been successfully expanded.

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.awt.tree.TreeExpansionEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="treeExpanded")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("treeExpanded", cb)

    def add_event_tree_expanding(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked whenever a node in the tree is about to be expanded.

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.awt.tree.TreeExpansionEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="treeExpanding")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("treeExpanding", cb)

    def add_event_tree_expansion_events_disposing(self, cb: EventArgsCallbackT) -> None:
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

    def remove_event_request_child_nodes(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="requestChildNodes", is_add=False)
            self.__callback(self, args)
        self.__listener.off("requestChildNodes", cb)

    def remove_event_tree_collapsed(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="treeCollapsed", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("treeCollapsed", cb)

    def remove_event_tree_collapsing(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="treeCollapsing", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("treeCollapsing", cb)

    def remove_event_tree_expanded(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="treeExpanded", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("treeExpanded", cb)

    def remove_event_tree_expanding(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="treeExpanding", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("treeExpanding", cb)

    def remove_event_tree_expansion_events_disposing(self, cb: EventArgsCallbackT) -> None:
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
    def events_listener_tree_expansion(self) -> TreeExpansionListener:
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
        source (Any): Expected to be an instance of TreeExpansionEvents that is a partial class of a component based class.
        event (ListenerEventArgs): Event arguments.

    Returns:
        None:

    Warning:
        This method is intended for internal use only.
    """
    # will only ever fire once
    if not isinstance(source, TreeExpansionEvents):
        return
    if not hasattr(source, "component"):
        return

    comp = cast("XTreeControl", source.component)  # type: ignore
    comp.addTreeExpansionListener(source.events_listener_tree_expansion)
    event.remove_callback = True
