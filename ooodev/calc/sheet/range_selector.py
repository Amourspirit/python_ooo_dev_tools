"""
Range selection dialog for Calc sheets.

Example Usage:

.. code-block:: python

    def main():
        with Lo.Loader(connector=Lo.ConnectSocket()) as loader:
            doc = None
            try:

                doc = CalcDoc.create_doc(loader=loader, visible=True)
                selector = RangeSelector()
                rng = selector.get_range_selection(doc)

                print("Range Sel", rng) # D3:E6

            finally:
                if doc is not None:
                    doc.close()
"""

from __future__ import annotations
import time
import contextlib
from typing import Any, cast, TYPE_CHECKING, Callable

import uno
import unohelper
from com.sun.star.sheet import XRangeSelectionListener
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.events.args.event_args import EventArgs
from ooodev.utils.helper.dot_dict import DotDict
from ooodev.utils.props import Props

if TYPE_CHECKING:
    from ooodev.calc import CalcSheetView
    from ooodev.utils.data_type.range_obj import RangeObj
    from com.sun.star.lang import EventObject
    from com.sun.star.sheet import RangeSelectionEvent
    from ooodev.calc import CalcDoc

# https://ask.libreoffice.org/t/range-selection-with-a-dialog-box-in-a-python-macro/33732/11


class RangeSelector:
    """
    Class to popup a range selector.

    Note:
        This class requires the GUI to be present and will not work in Headless mode.

        Also this class in implemented into ``CalcDoc`` and ``CalcSheet`` via the ``get_range_selection_from_popup()`` method.

    Example:

        .. code-block:: python

            doc = CalcDoc.from_current_doc()
            selector = RangeSelector()
            rng = selector.get_range_selection(doc)

            print("Range Sel", rng) # D3:E6


    .. versionadded:: 0.47.1
    """

    class _ExampleRangeListener(XRangeSelectionListener, EventsPartial, unohelper.Base):
        def __init__(self, view: CalcSheetView, auto_remove_listener: bool = True):
            XRangeSelectionListener.__init__(self)
            EventsPartial.__init__(self)
            unohelper.Base.__init__(self)
            self.view = view
            self.auto_remove = auto_remove_listener
            self._removed = False

        def done(self, event: RangeSelectionEvent):
            if event.RangeDescriptor:
                dd = DotDict(
                    state="done",
                    result=event.RangeDescriptor,
                    event=event,
                    view=self.view,
                    auto_remove=self.auto_remove,
                    rng_obj=None,
                )
            else:
                dd = DotDict(
                    state="aborted", result="", event=event, view=self.view, auto_remove=self.auto_remove, rng_obj=None
                )
            if dd.result:
                with contextlib.suppress(Exception):
                    sheet = self.view.calc_doc.get_active_sheet()
                    range = sheet.component.getCellRangeByName(dd.result)
                    dd.rng_obj = self.view.calc_doc.range_converter.get_range_obj(range)
            eargs = EventArgs(self)
            eargs.event_data = dd
            if self.auto_remove and self._removed is False:
                self.view.remove_range_selection_listener(self)
                self._removed = True
            self.trigger_event("RangeSelectResult", eargs)

        def aborted(self, event: RangeSelectionEvent):
            eargs = EventArgs(self)
            dd = DotDict(
                state="aborted",
                result="aborted",
                event=event,
                view=self.view,
                auto_remove=self.auto_remove,
                rng_obj=None,
            )
            eargs.event_data = dd
            if self.auto_remove and self._removed is False:
                self.view.remove_range_selection_listener(self)
                self._removed = True
            self.trigger_event("RangeSelectResult", eargs)

        def disposing(self, event: EventObject):
            pass

        def subscribe_range_select(self, cb: Callable[[Any, Any], None]) -> None:
            self.subscribe_event("RangeSelectResult", cb)

        def unsubscribe_range_select(self, cb: Callable[[Any, Any], None]) -> None:
            self.unsubscribe_event("RangeSelectResult", cb)

    def __init__(self, title: str = "Please select a range", close_on_mouse_release: bool = False):
        """
        Constructor for RangeSelection.

        Args:
            title (str, optional): The title of the popup. Defaults to "Please select a range".
            close_on_mouse_release (bool, optional): Specifies if the dialog closes when mouse is released. Defaults to ``False``.
        """
        self._title = title
        self._close_on_mouse_release = close_on_mouse_release
        self._init_cb()

    def _init_cb(self) -> None:
        self._fn_on_range_sel = self._on_range_sel

    def _on_range_sel(self, src: Any, event: EventArgs):
        view = cast("CalcSheetView", event.event_data.view)
        view.extra_data.selection_made = True
        # view.remove_range_selection_listener(src)
        if event.event_data.state == "done":
            if event.event_data.result:
                # print(f"Selected range: {event.event_data.result}")
                if event.event_data.rng_obj:
                    rng_obj = cast("RangeObj", event.event_data.rng_obj)
                    view.extra_data.selection_result = rng_obj

    def get_range_selection(self, doc: CalcDoc) -> RangeObj | None:
        """
        Get the range selection.

        Args:
            doc (CalcDoc): The CalcDoc object.

        Returns:
            RangeObj | None: The range object or ``None`` if no selection was made.
        """
        view = doc.get_view()
        ex_listener = RangeSelector._ExampleRangeListener(view)
        ex_listener.subscribe_range_select(self._fn_on_range_sel)
        view.extra_data.selection_made = False
        view.extra_data.selection_result = None
        view.add_range_selection_listener(ex_listener)
        props = Props.make_props(Title=self._title, CloseOnMouseRelease=self._close_on_mouse_release)
        view.component.startRangeSelection(props)
        # print("Make a selection in the document")
        while not view.extra_data.selection_made:
            time.sleep(0.5)
        result = view.extra_data.selection_result
        del view.extra_data["selection_made"]
        del view.extra_data["selection_result"]
        return result
