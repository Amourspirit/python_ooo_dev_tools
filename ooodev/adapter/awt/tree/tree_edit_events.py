from __future__ import annotations
from typing import TYPE_CHECKING

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
