from __future__ import annotations
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.events.lo_events import Events


class GblEvents(EventsPartial):
    """
    Singleton class for global events.

    Implements ``ooodev.events.events_t.EventsT`` protocol.

    This class is a singleton and is shared across all instances of the application.

    The purpose of this class to provide a global event system that can be used to trigger events across the application.

    .. versionadded:: 0.41.0
    """

    _instance = None

    def __new__(cls, *args, **kwargs):
        if not isinstance(cls._instance, cls):
            cls._instance = super(GblEvents, cls).__new__(cls, *args, **kwargs)
            cls._instance._initialized = False
        return cls._instance

    def __init__(self):
        """
        Construct for Global Events
        """
        if self._initialized:
            return
        super().__init__(Events(self))
        self._initialized = True
