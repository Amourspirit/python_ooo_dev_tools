from __future__ import annotations

from .tree_edit_listener import TreeEditListener
from .tree_edit_listener import NodeEditedArgs as NodeEditedArgs
from ooodev.adapter.adapter_base import GenericArgs
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils import gen_util as gUtil
from ooodev.utils.type_var import EventArgsCallbackT, ListenerEventCallbackT


class NodeEditEvents:
    """
    Class for managing Node Edit Events.

    This class is usually inherited by control classes that implement ``com.sun.star.awt.tree.XTreeEditListener``.
    """

    def __init__(self, trigger_args: GenericArgs | None = None, cb: ListenerEventCallbackT | None = None) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
            cb (ListenerEventCallbackT | None, optional): Callback that is invoked when an event is added or removed.
        """
        self.__callback = cb
        self.__listener = TreeEditListener(trigger_args=trigger_args)
        self.__name = gUtil.Util.generate_random_string(10)

    # region Manage Events
    def add_event_node_edited(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked from the TreeControl implementation when editing of Node is finished and was not canceled.

        Note:
            The callback ``EventArgs.event_data`` will contain a :py:class:`~ooodev.adapter.awt.tab.tree_edit_listener.NodeEditedArgs` object.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="nodeEdited")
            self.__callback(self, args)
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
        self.__listener.on("nodeEditing", cb)

    def remove_event_node_edited(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="nodeEdited", is_add=False)
            self.__callback(self, args)
        self.__listener.off("nodeEdited", cb)

    def remove_event_node_editing(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="nodeEditing", is_add=False)
            self.__callback(self, args)
        self.__listener.off("nodeEditing", cb)

    @property
    def events_listener_tab_page_container(self) -> TreeEditListener:
        """
        Returns listener
        """
        return self.__listener

    # endregion Manage Events
