from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING

from ooodev.events.args.generic_args import GenericArgs
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils import gen_util as gUtil
from ooodev.adapter.script.vba.vba_script_listener import VBAScriptListener

if TYPE_CHECKING:
    from com.sun.star.script.vba import XVBACompatibility
    from ooodev.utils.type_var import EventArgsCallbackT, ListenerEventCallbackT


class VBAScriptEvents:
    """
    Class for managing VBA Script Events.
    """

    def __init__(
        self,
        trigger_args: GenericArgs | None = None,
        cb: ListenerEventCallbackT | None = None,
        listener: VBAScriptListener | None = None,
        subscriber: XVBACompatibility | None = None,
    ) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
                This only applies if the listener is not passed.
            cb (ListenerEventCallbackT | None, optional): Callback that is invoked when an event is added or removed.
            listener (VBAScriptListener | None, optional): Listener that is used to manage events.
            subscriber (XVBACompatibility, optional): An UNO object that implements the ``XVBACompatibility`` interface.
                If passed in then this instance listener is automatically added to it.
        """
        self.__callback = cb
        if listener:
            self.__listener = listener
            if subscriber:
                subscriber.addVBAScriptListener(self.__listener)
        else:
            self.__listener = VBAScriptListener(trigger_args=trigger_args, subscriber=subscriber)
        self.__name = gUtil.Util.generate_random_string(10)

    # region Manage Events
    def add_event_notify_vba_script_event(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when the object is about to be reloaded.

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.script.vba.VBAScriptEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="notifyVBAScriptEvent")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("notifyVBAScriptEvent", cb)

    def add_event_vba_script_events_disposing(self, cb: EventArgsCallbackT) -> None:
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

    def remove_event_notify_vba_script_event(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event.
        """
        # renamed in version 0.19.1, was named remove_event_document_event_occured, alias added below
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="notifyVBAScriptEvent", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("notifyVBAScriptEvent", cb)

    def remove_event_vba_script_events_disposing(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="disposing", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("disposing", cb)

    @property
    def events_listener_vba_script(self) -> VBAScriptListener:
        """
        Returns listener.
        """
        return self.__listener

    # endregion Manage Events


def on_lazy_cb(source: Any, event: ListenerEventArgs) -> None:
    """
    Callback that is invoked when an event is added or removed.

    This method is generally used to add the listener to the component in a lazy manner.
    This means this callback will only be called once in the lifetime of the component.

    Args:
        source (Any): Expected to be an instance of VBAScriptEvents that is a partial class of a component based class.
        event (ListenerEventArgs): Event arguments.

    Returns:
        None:

    Warning:
        This method is intended for internal use only.
    """
    # will only ever fire once
    if not isinstance(source, VBAScriptEvents):
        return
    if not hasattr(source, "component"):
        return

    comp = cast("XVBACompatibility", source.component)  # type: ignore
    comp.addVBAScriptListener(source.events_listener_vba_script)
    event.remove_callback = True
