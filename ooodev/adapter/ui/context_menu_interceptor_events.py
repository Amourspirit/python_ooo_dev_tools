from __future__ import annotations
from typing import Any, cast, Dict, TYPE_CHECKING

from ooodev.events.args.generic_args import GenericArgs
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils import gen_util as gUtil
from ooodev.adapter.ui.context_menu_interceptor import ContextMenuInterceptor

if TYPE_CHECKING:
    from com.sun.star.ui import XContextMenuInterception
    from ooodev.utils.type_var import EventArgsCallbackT, ListenerEventCallbackT


# https://wiki.documentfoundation.org/Framework/Tutorial/Context_Menu_Interception


class ContextMenuInterceptorEvents:
    """
    Class for managing Context Menu Interceptor Events.
    """

    def __init__(
        self,
        component: XContextMenuInterception,
        trigger_args: GenericArgs | None = None,
        cb: ListenerEventCallbackT | None = None,
    ) -> None:
        """
        Constructor

        Args:
            component (XContextMenuInterception): UNO Component that implements ``com.sun.star.ui.XContextMenuInterception``.
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
                This only applies if the listener is not passed.
            cb (ListenerEventCallbackT | None, optional): Callback that is invoked when an event is added or removed.
        """
        self.__component = component
        self.__callback = cb
        self.__trigger_args = trigger_args
        self.__events: Dict[EventArgsCallbackT, ContextMenuInterceptor] = {}
        self.__name = gUtil.Util.generate_random_string(10)

    # region Manage Events
    def add_event_notify_context_menu_execute(self, cb: EventArgsCallbackT) -> None:
        """
        Adds and registers a listener for an event.

        Notifies the interceptor about the request to execute a ContextMenu.
        The interceptor has to decide whether the menu should be executed with or without being modified or may ignore the call.

        The call back ``event`` is an instance of ``GenericArgs[ContextMenuInterceptorEventData]``.

        The callback ``GenericArgs.event_data`` will be an instance of :py:class:`~ooodev.adapter.ui.context_menu_interceptor_event_data.ContextMenuInterceptorEventData`.

        The ``event_data.action`` defaults to ``ContextMenuInterceptorAction.IGNORED`` and can be modified to change the action.
        ``IGNORED`` in this context means that The next registered ``XContextMenuInterceptor`` should be notified.

        Returns:
            None:

        Hint:
        - ``ContextMenuInterceptorAction`` is an enum that can be imported from ``ooo.dyn.ui.context_menu_interceptor_action``.

        Note:
            Registering an interceptor adds it to the front of the interceptor chain, so that it is called first.
            The order of removals is arbitrary. It is not necessary to remove the interceptor that registered last.

            Each time this method is called, a new listener is created and registered if it does not already exist.

        See Also:
            - :py:class:`~ooodev.adapter.ui.context_menu_interceptor_event_data.ContextMenuInterceptorEventData`
            - :py:class:`~ooodev.events.args.generic_args.GenericArgs`
            - `Context Menu Interception <https://wiki.documentfoundation.org/Framework/Tutorial/Context_Menu_Interception>`__
        """

        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="notifyContextMenuExecute")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        if cb in self.__events:
            return
        listener = ContextMenuInterceptor(trigger_args=self.__trigger_args)
        listener.on("notifyContextMenuExecute", cb)
        self.__events[cb] = listener
        self.__component.registerContextMenuInterceptor(listener)

    def remove_event_notify_context_menu_execute(self, cb: EventArgsCallbackT) -> None:
        """
        Un-registers and removes a listener for an event.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="notifyContextMenuExecute", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        if cb not in self.__events:
            return
        listener = self.__events.pop(cb)
        listener.off("notifyContextMenuExecute", cb)
        self.__component.releaseContextMenuInterceptor(listener)

    # endregion Manage Events


def on_lazy_cb(source: Any, event: ListenerEventArgs) -> None:
    """
    Callback that is invoked when an event is added or removed.

    This method is generally used to add the listener to the component in a lazy manner.
    This means this callback will only be called once in the lifetime of the component.

    Args:
        source (Any): Expected to be an instance of ActivationEventEvents that is a partial class of a component based class.
        event (ListenerEventArgs): Event arguments.

    Returns:
        None:

    Warning:
        This method is intended for internal use only.
    """
    # will only ever fire once
    if not isinstance(source, ContextMenuInterceptorEvents):
        return
    # if not hasattr(source, "component"):
    #     return

    # There is not attaching events to the component for ContextMenuInterceptorEvents.
    event.remove_callback = True
