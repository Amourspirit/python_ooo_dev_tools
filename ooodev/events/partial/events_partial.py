from __future__ import annotations
from typing import cast, TYPE_CHECKING
from ooodev.events.args.event_args_t import EventArgsT
from ooodev.events.lo_events import Events
from ooodev.utils.type_var import EventCallback

if TYPE_CHECKING:
    from ooodev.proto.event_observer import EventObserver


class EventsPartial:
    def __init__(self, events: EventObserver | None = None):
        if events is not None:
            self.__events = events
        else:
            self.__events = cast("EventObserver", Events(source=self))

    # region Events

    def add_event_observers(self, *args: EventObserver) -> None:
        """
        Adds observers that gets their ``trigger`` method called when this class ``trigger`` method is called.

        Parameters:
            args (EventObserver): One or more observers to add.

        Returns:
            None:

        Note:
            Observers are removed automatically when they are out of scope.
        """
        self.__events.add_observer(*args)

    def remove_event_observer(self, observer: EventObserver) -> bool:
        """
        Removes an observer

        Args:
            observer (EventObserver): One or more observers to add.

        Returns:
            bool: ``True`` if observer has been removed; Otherwise, ``False``.
        """
        return self.__events.remove_observer(observer)

    def subscribe_event(self, event_name: str, callback: EventCallback) -> None:
        """
        Add an event listener to current instance.

        Args:
            event_name (str): Event Name.
            callback (EventCallback): Callback of the event listener.

        Returns:
            None:
        """
        self.__events.on(event_name, callback)

    def unsubscribe_event(self, event_name: str, callback: EventCallback) -> None:
        """
        Remove an event listener from current instance.

        Args:
            event_name (str): Event Name.
            callback (EventCallback): Callback of the event listener.

        Returns:
            None:
        """
        self.__events.remove(event_name, callback)

    def trigger_event(self, event_name: str, event_args: EventArgsT):
        """
        Trigger an event on current instance.

        Args:
            event_name (str): Event Name.
            event_args (EventArgsT): Event Args.

        Returns:
            None:
        """
        self.__events.trigger(event_name, event_args)

    # endregion Events

    @property
    def event_observer(self) -> EventObserver:
        """Gets/Sets The Event Observer for this instance."""
        return self.__events

    @event_observer.setter
    def event_observer(self, value: EventObserver) -> None:
        """Sets The Event Observer for this instance."""
        self.__events = value
