from __future__ import annotations
from typing import Any, TYPE_CHECKING

if TYPE_CHECKING:
    from typing_extensions import Protocol
    from ooodev.events.args.event_args_t import EventArgsT
    from ooodev.utils.type_var import EventCallback
    from ooodev.proto.event_observer import EventObserver
else:
    Protocol = object
    EventArgsT = Any
    EventCallback = Any
    EventObserver = Any


class EventsT(Protocol):
    """
    Protocol Class for Events.
    """

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
        ...

    def remove_event_observer(self, observer: EventObserver) -> bool:
        """
        Removes an observer

        Args:
            observer (EventObserver): One or more observers to add.

        Returns:
            bool: ``True`` if observer has been removed; Otherwise, ``False``.
        """
        ...

    def subscribe_event(self, event_name: str, callback: EventCallback) -> None:
        """
        Add an event listener to current instance.

        Args:
            event_name (str): Event Name.
            callback (EventCallback): Callback of the event listener.

        Returns:
            None:
        """
        ...

    def unsubscribe_event(self, event_name: str, callback: EventCallback) -> None:
        """
        Remove an event listener from current instance.

        Args:
            event_name (str): Event Name.
            callback (EventCallback): Callback of the event listener.

        Returns:
            None:
        """
        ...

    def trigger_event(self, event_name: str, event_args: EventArgsT):
        """
        Trigger an event on current instance.

        Args:
            event_name (str): Event Name.
            event_args (EventArgsT): Event Args.

        Returns:
            None:
        """
        ...

    @property
    def event_observer(self) -> EventObserver:
        """Gets/Sets The Event Observer for this instance."""
        ...

    @event_observer.setter
    def event_observer(self, value: EventObserver) -> None:
        """Sets The Event Observer for this instance."""
        ...
