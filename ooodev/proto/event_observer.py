from __future__ import annotations
from typing import Any, TYPE_CHECKING

if TYPE_CHECKING:
    from typing_extensions import Protocol
    from ooodev.events.args.event_args_t import EventArgsT
    from ooodev.utils.type_var import EventCallback
else:
    Protocol = object
    EventArgsT = Any
    EventCallback = Any


class EventObserver(Protocol):
    """
    Protocol Class for Event Observer.

    See Also:
        :py:mod:`~.events.lo_events`
    """

    def add_observer(self, *args: EventObserver) -> None:
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

    def remove_observer(self, observer: EventObserver) -> bool:
        """
        Removes an observer

        Args:
            observer (EventObserver): One or more observers to add.

        Returns:
            bool: ``True`` if observer has been removed; Otherwise, ``False``.
        """
        ...

    def on(self, event_name: str, callback: EventCallback):
        """
        Registers an event

        Args:
            event_name (str): Unique event name
            callback (Callable[[object, EventArgs], None]): Callback function
        """
        ...

    def remove(self, event_name: str, callback: EventCallback) -> bool:
        """
        Removes an event callback

        Args:
            event_name (str): Unique event name
            callback (Callable[[object, EventArgs], None]): Callback function

        Returns:
            bool: True if callback has been removed; Otherwise, False.
            False means the callback was not found.
        """
        ...

    def trigger(self, event_name: str, event_args: EventArgsT):
        """
        Trigger event(s) for a given name.

        Args:
            event_name (str): Name of event to trigger
            event_args (EventArgsT): Event args passed to the callback for trigger.
        """
        ...
