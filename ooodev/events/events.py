from __future__ import annotations
from typing import Any
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.events.lo_events import Events as LoEvents
from ooodev.events.args.generic_args import GenericArgs


class Events(EventsPartial):
    """
    Basic class for events.

    Implements ``ooodev.events.events_t.EventsT`` protocol.

    .. versionadded:: 0.32.0
    """

    def __init__(self, source: Any = None, trigger_args: GenericArgs | None = None):
        """
        Construct for Events

        Args:
            source (Any | None, optional): Source can be class or any object.
                The value of ``source`` is the value assigned to the ``EventArgs.event_source`` property.
                Defaults to current instance of this class.
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
        """
        if source is None:
            source = self
        super().__init__(LoEvents(source, trigger_args))
