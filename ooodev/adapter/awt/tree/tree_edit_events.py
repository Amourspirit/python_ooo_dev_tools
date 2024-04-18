from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING

from ooodev.events.args.generic_args import GenericArgs
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils import gen_util as gUtil
from ooodev.utils.type_var import EventArgsCallbackT, ListenerEventCallbackT
from ooodev.adapter.awt.tree.tree_edit_listener import TreeEditListener

# pylint: disable=useless-import-alias, unused-import
from ooodev.adapter.awt.tree.tree_edit_listener import NodeEditedArgs as NodeEditedArgs

if TYPE_CHECKING:
    from com.sun.star.awt.tree import XTreeControl


class TreeEditEvents:
    """
    Class for managing Tree Node Edit Events.

    This class is usually inherited by control classes that implement ``com.sun.star.awt.tree.XTreeEditListener``.
    """

    def __init__(
        self,
        trigger_args: GenericArgs | None = None,
        cb: ListenerEventCallbackT | None = None,
        listener: TreeEditListener | None = None,
        subscriber: XTreeControl | None = None,
    ) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
                This only applies if the listener is not passed.
            cb (ListenerEventCallbackT | None, optional): Callback that is invoked when an event is added or removed.
            listener (TreeEditListener | None, optional): Listener that is used to manage events.
            subscriber (XTreeControl, optional): An UNO object that implements the ``XTreeControl`` interface.
                If passed in then this instance listener is automatically added to it.
        """
        self.__callback = cb
        if listener:
            self.__listener = listener
            if subscriber:
                subscriber.addTreeEditListener(self.__listener)
        else:
            self.__listener = TreeEditListener(trigger_args=trigger_args, subscriber=subscriber)
        self.__name = gUtil.Util.generate_random_string(10)

    # region Manage Events
    def add_event_node_edited(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked from the TreeControl implementation when editing of Node is finished and was not canceled.

        Note:
            The callback ``EventArgs.event_data`` will contain a :py:class:`~ooodev.adapter.awt.tab.tree_edit_listener.NodeEditedArgs` object.
        """
        # sourcery skip: class-extract-method
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="nodeEdited")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("nodeEdited", cb)

    def add_event_node_editing(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked after a tab page was activated.

        The callback ``EventArgs.event_data`` will contain a UNO object that implements ``XTreeNode``.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="nodeEditing")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("nodeEditing", cb)

    def add_event_tree_edit_events_disposing(self, cb: EventArgsCallbackT) -> None:
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

    def remove_event_node_edited(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="nodeEdited", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("nodeEdited", cb)

    def remove_event_node_editing(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="nodeEditing", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("nodeEditing", cb)

    def remove_event_tree_edit_events_disposing(self, cb: EventArgsCallbackT) -> None:
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
    def events_listener_tree_edit(self) -> TreeEditListener:
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
        source (Any): Expected to be an instance of TreeEditEvents that is a partial class of a component based class.
        event (ListenerEventArgs): Event arguments.

    Returns:
        None:

    Warning:
        This method is intended for internal use only.
    """
    # will only ever fire once
    if not isinstance(source, TreeEditEvents):
        return
    if not hasattr(source, "component"):
        return

    comp = cast("XTreeControl", source.component)  # type: ignore
    comp.addTreeEditListener(source.events_listener_tree_edit)
    event.remove_callback = True
