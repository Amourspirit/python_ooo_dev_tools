from __future__ import annotations
from typing import TYPE_CHECKING
import uno
from com.sun.star.util import XCloseBroadcaster

from ooodev.events.args.generic_args import GenericArgs
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils import gen_util as gUtil
from ooodev.adapter.util.close_listener import CloseListener

if TYPE_CHECKING:
    from ooodev.utils.type_var import EventArgsCallbackT, ListenerEventCallbackT


class CloseEvents:
    """
    Class for managing Close Events.

    This class is usually inherited by control classes that implement ``com.sun.star.util.XCloseListener``.

    .. versionadded:: 0.21.0
    """

    def __init__(
        self,
        trigger_args: GenericArgs | None = None,
        cb: ListenerEventCallbackT | None = None,
        listener: CloseListener | None = None,
        subscriber: XCloseBroadcaster | None = None,
    ) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
                This only applies if the listener is not passed.
            cb (ListenerEventCallbackT | None, optional): Callback that is invoked when an event is added or removed.
            listener (CloseListener | None, optional): Listener that is used to manage events.
            subscriber (XCloseBroadcaster, optional): An UNO object that implements the ``XCloseBroadcaster`` interface.
                If passed in then this instance listener is automatically added to it.
        """
        self.__callback = cb
        if listener:
            self.__listener = listener
            if subscriber:
                subscriber.addCloseListener(self.__listener)
        else:
            self.__listener = CloseListener(trigger_args=trigger_args, subscriber=subscriber)
        self.__name = gUtil.Util.generate_random_string(10)

    # region Manage Events
    def add_event_notify_closing(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when the listened object is closed really.

        Now the listened object is closed really.
        Listener has to accept that; should deregister itself and release all references to it.
        It's not allowed nor possible to disagree with that by throwing any exception.

        If the event ``com.sun.star.lang.XEventListener.disposing()`` occurred before it must be accepted too.
        There exist no chance for a disagreement any more.

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.lang.EventObject`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="notifyClosing")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("notifyClosing", cb)

    def add_event_query_closing(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when somewhere tries to close listened object

        Is called before ``XCloseListener.notifyClosing()``.
        Listener has the chance to break that by throwing a CloseVetoException.
        This exception must be passed to the original caller of ``XCloseable.close()`` without any interaction.

        The parameter ``gets_ownership`` regulate who has to try to close the listened object again,
        if this listener disagree with the request by throwing the exception.
        If it's set to ``False`` the original caller of ``XCloseable.close()`` will be the owner in every case.
        It's not allowed to call ``close()`` from this listener then.
        If it's set to ``True`` this listener will be the new owner if he throw the exception, otherwise not!
        If his still running processes will be finished he must call ``close()`` on listened object again then.

        If this listener doesn't disagree with th close request it depends from his internal implementation if he deregister itself at the listened object.
        But normally this must be done in ``XCloseListener.notifyClosing()``.

        The callback ``KeyValArgs.event_data`` will contain a UNO ``com.sun.star.lang.EventObject`` struct.

        Note:
            The callback event is an instance of :py:class:`~ooodev.events.args.key_val_args.KeyValArgs` with the following properties:
            ``event.key=gets_ownership``, ``event.value`` is a bool and ``event.event_data`` is ``com.sun.star.lang.EventObject``.

            This is because the UNO interface ``com.sun.star.util.XCloseListener`` has the following signature:
            ``void queryClosing	([in] com::sun::star::lang::EventObject Source, [in] boolean GetsOwnership)``.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="queryClosing")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("queryClosing", cb)

    def add_event_close_events_disposing(self, cb: EventArgsCallbackT) -> None:
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

    def remove_event_notify_closing(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="notifyClosing", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("notifyClosing", cb)

    def remove_event_query_closing(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="queryClosing", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("queryClosing", cb)

    def remove_event_close_events_disposing(self, cb: EventArgsCallbackT) -> None:
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
    def events_listener_close(self) -> CloseListener:
        """
        Returns listener
        """
        return self.__listener

    # endregion Manage Events
