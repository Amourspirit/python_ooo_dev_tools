from __future__ import annotations
from typing import Any

# pylint: disable=unused-import, useless-import-alias
import uno
from ..events.args.event_args import EventArgs as EventArgs
from ..events.lo_events import Events, EventCallback
from ..events.args.generic_args import GenericArgs as GenericArgs
from ..mock import mock_g

if mock_g.DOCS_BUILDING:
    from ..mock import unohelper
else:
    import unohelper


class AdapterBase(unohelper.Base):  # type: ignore
    """
    Base Class for Listeners in the ``adapter`` name space.
    """

    def __init__(self, trigger_args: GenericArgs | None) -> None:
        """
        Constructor

        Arguments:
            trigger_args (GenericArgs, Optional): Args that are passed to events when they are triggered.
        """
        super().__init__()
        self._events = Events(source=self, trigger_args=trigger_args)

    def _trigger_event(self, name: str, event: Any) -> None:
        # any trigger args passed in will be passed to callback event via Events class.
        event_arg = EventArgs(self.__class__.__qualname__)
        event_arg.event_data = event
        self._events.trigger(name, event_arg)

    def on(self, event_name: str, cb: EventCallback) -> None:
        """
        Adds a listener for an event

        Args:
            event_name (str): Event name to add listener for. Usually the name of the method being listened to such as ``windowOpened``
            cb (EventCallback): Callback event.
        """
        self._events.on(event_name, cb)

    def off(self, event_name: str, cb: EventCallback) -> None:
        """
        Removes a listener for an event

        Args:
            event_name (str): Event Name
            cb (EventCallback): Callback event
        """
        self._events.remove(event_name, cb)

    def clear(self) -> None:
        """
        Removes all listeners for all events

        .. versionadded:: 0.13.7
        """
        self._events.clear()
