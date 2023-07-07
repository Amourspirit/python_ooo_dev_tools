from __future__ import annotations
from com.sun.star.frame import XComponentLoader
from ooodev.utils.lo import Lo
from ooodev.events.event_singleton import _Events
from ooodev.events.lo_named_event import LoNamedEvent
from ooodev.events.args.event_args import EventArgs


class MacroLoader:
    """
    Context Manager for Macro execution

    .. versionadded:: 0.11.11
    """

    def __init__(self):
        """
        Create a connection to running instance of office
        """
        self.loader = Lo.load_office()

    def __enter__(self) -> XComponentLoader:
        _Events().trigger(LoNamedEvent.MACRO_LOADER_ENTER, EventArgs("MacroLoader.__enter__"))
        return self.loader

    def __exit__(self, exc_type, exc_val, exc_tb):
        _Events().trigger(LoNamedEvent.MACRO_LOADER_EXIT, EventArgs("MacroLoader.__exit__"))
        # Lo.close_office()
        return
