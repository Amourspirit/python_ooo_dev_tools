from __future__ import annotations
from ooodev.events.args.event_args_t import EventArgsT
from ooodev.events.lo_events import Events
from ooodev.utils.type_var import EventCallback


class EventsPartial:
    def __init__(self):
        self.__events = Events(source=self)

    # region Events

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
