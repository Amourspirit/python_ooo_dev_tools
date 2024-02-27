from __future__ import annotations

from typing import TYPE_CHECKING

from ooodev.events.args.generic_args import GenericArgs
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils import gen_util as gUtil
from ooodev.adapter.form.load_listener import LoadListener

if TYPE_CHECKING:
    from com.sun.star.form import XLoadable
    from ooodev.utils.type_var import EventArgsCallbackT, ListenerEventCallbackT


class LoadEvents:
    """
    Class for managing Load Events.
    """

    def __init__(
        self,
        trigger_args: GenericArgs | None = None,
        cb: ListenerEventCallbackT | None = None,
        listener: LoadListener | None = None,
        subscriber: XLoadable | None = None,
    ) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
                This only applies if the listener is not passed.
            cb (ListenerEventCallbackT | None, optional): Callback that is invoked when an event is added or removed.
            listener (LoadListener | None, optional): Listener that is used to manage events.
            subscriber (XLoadable, optional): An UNO object that implements the ``com.sun.star.form.XLoadable`` interface.
                If passed in then this instance listener is automatically added to it.
        """
        self.__callback = cb
        if listener:
            self.__listener = listener
            if subscriber:
                subscriber.addLoadListener(self.__listener)
        else:
            self.__listener = LoadListener(trigger_args=trigger_args, subscriber=subscriber)
        self.__name = gUtil.Util.generate_random_string(10)

    def __add_listener(self, trigger_name: str, cb: EventArgsCallbackT) -> None:
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name=trigger_name)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on(trigger_name, cb)

    def __remove_listener(self, trigger_name: str, cb: EventArgsCallbackT) -> None:
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name=trigger_name, is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off(trigger_name, cb)

    # region Manage Events
    def add_event_loaded(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when the object has successfully connected to a datasource.

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.lang.EventObject`` struct.
        """
        self.__add_listener("loaded", cb)

    def add_event_reloaded(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when the object has been reloaded.

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.lang.EventObject`` struct.
        """
        self.__add_listener("reloaded", cb)

    def add_event_reloading(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when the object is about to be reloaded.

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.lang.EventObject`` struct.
        """
        self.__add_listener("reloading", cb)

    def add_event_unloaded(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked after the object has disconnected from a datasource.

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.lang.EventObject`` struct.
        """
        self.__add_listener("unloaded", cb)

    def add_event_unloading(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when the object is about to be unloaded.

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.lang.EventObject`` struct.
        """
        self.__add_listener("unloading", cb)

    def add_event_load_events_events_disposing(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when the broadcaster is about to be disposed.

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.lang.EventObject`` struct.
        """
        self.__add_listener("disposing", cb)

    def remove_event_loaded(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        self.__remove_listener("loaded", cb)

    def remove_event_reloaded(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        self.__remove_listener("reloaded", cb)

    def remove_event_reloading(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        self.__remove_listener("reloading", cb)

    def remove_event_unloaded(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        self.__remove_listener("unloaded", cb)

    def remove_event_unloading(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        self.__remove_listener("unloading", cb)

    def remove_event_load_events_disposing(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        self.__remove_listener("disposing", cb)

    @property
    def events_listener_load(self) -> LoadListener:
        """
        Returns listener
        """
        return self.__listener

    # endregion Manage Events
