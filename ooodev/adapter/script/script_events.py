from __future__ import annotations

from typing import TYPE_CHECKING

from ooodev.events.args.generic_args import GenericArgs
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils import gen_util as gUtil
from ooodev.adapter.script.script_listener import ScriptListener

if TYPE_CHECKING:
    from com.sun.star.script import XEventAttacherManager
    from ooodev.utils.type_var import EventArgsCallbackT, CancelEventArgsCallbackT, ListenerEventCallbackT


class ScriptEvents:
    """
    Class for managing Script Events.
    """

    def __init__(
        self,
        trigger_args: GenericArgs | None = None,
        cb: ListenerEventCallbackT | None = None,
        listener: ScriptListener | None = None,
        subscriber: XEventAttacherManager | None = None,
    ) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
                This only applies if the listener is not passed.
            cb (ListenerEventCallbackT | None, optional): Callback that is invoked when an event is added or removed.
            listener (ScriptListener | None, optional): Listener that is used to manage events.
            subscriber (XEventAttacherManager, optional): An UNO object that implements the ``com.sun.star.form.XEventAttacherManager`` interface.
                If passed in then this instance listener is automatically added to it.
        """
        self.__callback = cb
        if listener:
            self.__listener = listener
            if subscriber:
                subscriber.addScriptListener(self.__listener)
        else:
            self.__listener = ScriptListener(trigger_args=trigger_args, subscriber=subscriber)
        self.__name = gUtil.Util.generate_random_string(10)

    # region Manage Events
    def add_event_approve_firing(self, cb: CancelEventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked is invoked when an firing is performed.
        If event is canceled then the firing will be canceled.

        The callback ``CancelEventArgs.event_data`` will contain a UNO ``com.sun.star.script.ScriptEvent`` struct.

        Note:
            The callback event will be :py:class:`~ooodev.events.args.cancel_event_args.CancelEventArgs`.
            If the ``CancelEventArgs.cancel`` is set to ``True`` then the firing will be canceled if the ``CancelEventArgs.handled``
            is set to ``True`` then the firing will be performed.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="approveFiring")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("approveFiring", cb)

    def add_event_firing(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when an event takes place.

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.script.ScriptEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="firing")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("firing", cb)

    def add_event_script_events_disposing(self, cb: EventArgsCallbackT) -> None:
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

    def remove_event_approve_firing(self, cb: CancelEventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="approveFiring", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("approveFiring", cb)

    def remove_event_firing(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="firing", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("firing", cb)

    def remove_event_script_events_disposing(self, cb: EventArgsCallbackT) -> None:
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
    def events_listener_script(self) -> ScriptListener:
        """
        Returns listener
        """
        return self.__listener

    # endregion Manage Events
