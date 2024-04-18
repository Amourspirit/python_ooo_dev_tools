from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
from ooodev.events.args.generic_args import GenericArgs
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils import gen_util as gUtil
from ooodev.adapter.awt.menu_listener import MenuListener

if TYPE_CHECKING:
    from com.sun.star.awt import XMenu
    from ooodev.utils.type_var import EventArgsCallbackT, ListenerEventCallbackT


class MenuEvents:
    """
    Class for managing Menu Events.
    """

    def __init__(
        self,
        trigger_args: GenericArgs | None = None,
        cb: ListenerEventCallbackT | None = None,
        listener: MenuListener | None = None,
        subscriber: XMenu | None = None,
    ) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
                This only applies if the listener is not passed.
            cb (ListenerEventCallbackT | None, optional): Callback that is invoked when an event is added or removed.
            listener (MenuListener | None, optional): Listener that is used to manage events.
            subscriber (XMenu, optional): An UNO object that implements the ``XMenu`` interface.
                If passed in then this instance listener is automatically added to it.
        """
        self.__callback = cb
        if listener:
            self.__listener = listener
            if subscriber:
                subscriber.addMenuListener(self.__listener)
        else:
            self.__listener = MenuListener(trigger_args=trigger_args, subscriber=subscriber)
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

    def add_event_item_activated(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when a key has been pressed.

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.awt.MenuEvent`` struct.
        """
        self.__add_listener("itemActivated", cb)

    def add_event_item_deactivated(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when a key has been pressed.

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.awt.MenuEvent`` struct.
        """
        self.__add_listener("itemDeactivated", cb)

    def add_event_item_highlighted(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when a key has been pressed.

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.awt.MenuEvent`` struct.
        """
        self.__add_listener("itemHighlighted", cb)

    def add_event_item_selected(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when a key has been pressed.

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.awt.MenuEvent`` struct.
        """
        self.__add_listener("itemSelected", cb)

    def add_event_menu_events_disposing(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when the broadcaster is about to be disposed.

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.lang.EventObject`` struct.
        """
        self.__add_listener("disposing", cb)

    def remove_event_item_activated(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        self.__remove_listener("itemActivated", cb)

    def remove_event_item_deactivated(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        self.__remove_listener("itemDeactivated", cb)

    def remove_event_item_highlighted(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        self.__remove_listener("itemHighlighted", cb)

    def remove_event_item_selected(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        self.__remove_listener("itemSelected", cb)

    def remove_event_menu_events_disposing(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        self.__remove_listener("disposing", cb)

    # endregion Manage Events

    @property
    def events_listener_menu(self) -> MenuListener:
        """
        Returns listener
        """
        return self.__listener


def on_lazy_cb(source: Any, event: ListenerEventArgs) -> None:
    """
    Callback that is invoked when an event is added or removed.

    This method is generally used to add the listener to the component in a lazy manner.
    This means this callback will only be called once in the lifetime of the component.

    Args:
        source (Any): Expected to be an instance of MenuEvents that is a partial class of a component based class.
        event (ListenerEventArgs): Event arguments.

    Returns:
        None:

    Warning:
        This method is intended for internal use only.
    """
    # will only ever fire once
    if not isinstance(source, MenuEvents):
        return
    if not hasattr(source, "component"):
        return

    comp = cast("XMenu", source.component)  # type: ignore
    comp.addMenuListener(source.events_listener_menu)
    event.remove_callback = True
