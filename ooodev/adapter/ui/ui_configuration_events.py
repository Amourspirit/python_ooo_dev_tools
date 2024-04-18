from __future__ import annotations

from typing import Any, cast, TYPE_CHECKING
from ooodev.events.args.generic_args import GenericArgs
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils import gen_util as gUtil
from ooodev.adapter.ui.ui_configuration_listener import UIConfigurationListener

if TYPE_CHECKING:
    from ooodev.utils.type_var import EventArgsCallbackT, ListenerEventCallbackT
    from com.sun.star.ui import XUIConfiguration


class UIConfigurationEvents:
    """
    Class for managing UI Configuration Events.
    """

    def __init__(
        self,
        trigger_args: GenericArgs | None = None,
        cb: ListenerEventCallbackT | None = None,
        listener: UIConfigurationListener | None = None,
        subscriber: XUIConfiguration | None = None,
    ) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
                This only applies if the listener is not passed.
            cb (ListenerEventCallbackT | None, optional): Callback that is invoked when an event is added or removed.
            listener (UIConfigurationListener | None, optional): Listener that is used to manage events.
            subscriber (XUIConfiguration, optional): An UNO object that implements ``com.sun.star.ui.XUIConfiguration``.
                If passed in then this instance listener is automatically added to it.
        """
        self.__callback = cb
        if listener:
            self.__listener = listener
            if subscriber:
                subscriber.addConfigurationListener(self.__listener)
        else:
            self.__listener = UIConfigurationListener(trigger_args=trigger_args, subscriber=subscriber)
        self.__name = gUtil.Util.generate_random_string(10)

    # region Manage Events
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

    def add_event_element_inserted(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is when a configuration has inserted an user interface element.

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.ui.ConfigurationEvent`` struct.
        """
        self.__add_listener("elementInserted", cb)

    def add_event_element_removed(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when a configuration has removed an user interface element.

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.ui.ConfigurationEvent`` struct.
        """
        self.__add_listener("elementRemoved", cb)

    def add_event_element_replaced(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event invoked when a configuration has replaced an user interface element.

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.ui.ConfigurationEvent`` struct.
        """
        self.__add_listener("elementReplaced", cb)

    def add_event_ui_configuration_disposing(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when the broadcaster is about to be disposed.

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.lang.EventObject`` struct.
        """

        self.__remove_listener("disposing", cb)

    def remove_event_element_inserted(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        self.__remove_listener("elementInserted", cb)

    def remove_event_element_removed(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        self.__remove_listener("elementRemoved", cb)

    def remove_event_element_replaced(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        self.__remove_listener("elementReplaced", cb)

    def remove_event_ui_configuration_disposing(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        self.__add_listener("disposing", cb)

    @property
    def events_listener_ui_configuration(self) -> UIConfigurationListener:
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
        source (Any): Expected to be an instance of UIConfigurationEvents that is a partial class of a component based class.
        event (ListenerEventArgs): Event arguments.

    Returns:
        None:

    Warning:
        This method is intended for internal use only.
    """
    # will only ever fire once
    if not isinstance(source, UIConfigurationEvents):
        return
    if not hasattr(source, "component"):
        return

    comp = cast("XUIConfiguration", source.component)  # type: ignore
    comp.addConfigurationListener(source.events_listener_ui_configuration)
    event.remove_callback = True
