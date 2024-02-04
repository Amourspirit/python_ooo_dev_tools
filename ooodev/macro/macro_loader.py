from __future__ import annotations
import os
from com.sun.star.frame import XComponentLoader
from ooodev.loader.lo import Lo
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
        Context Manager for Macro execution

        This context manager is used in a running instance of office.

        The context manager can be overridden to return the current loader (``Lo.loader_current``) of the Lo class.
        To override the context manager, set the environment variable ``ODEV_MACRO_LOADER_OVERRIDE`` to ``1``.
        Overriding is useful for testing, debugging and when you want to run a macro script outside of a running instance of office.

        Example:
            .. code-block:: python

                from ooodev.macro.macro_loader import MacroLoader

                def show_tab_dialog(*args, **kwargs):
                    with MacroLoader():
                        dlg = MultiSyntaxController(model=MultiSyntaxModel(), view=MultiSyntaxView())
                        dlg.start()

        .. versionadded:: 0.11.11

        .. versionchanged:: 0.11.14
            Added override functionality
        """
        self._override = False
        override_loader = os.environ.get("ODEV_MACRO_LOADER_OVERRIDE", "")
        if override_loader == "1":
            try:
                self.loader = Lo.loader_current
                self._override = True
            except AttributeError:
                self.loader = Lo.load_office()
        else:
            self.loader = Lo.load_office()
        self._inst = Lo.current_lo

    def __enter__(self) -> XComponentLoader:
        if self._override:
            return self.loader
        _Events().trigger(LoNamedEvent.MACRO_LOADER_ENTER, EventArgs("MacroLoader.__enter__"))
        if self._inst is None:
            raise RuntimeError("No running instance of office")
        if self._inst.this_component is None:
            # if this_component ten load_component() has already been called.
            raise RuntimeError("Critical error: Unable to get ThisComponent")
        return self.loader

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self._override:
            return
        _Events().trigger(LoNamedEvent.MACRO_LOADER_EXIT, EventArgs("MacroLoader.__exit__"))
        # Lo.close_office()
        return
