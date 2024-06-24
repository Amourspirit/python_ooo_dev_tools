from __future__ import annotations
from typing import Any, TYPE_CHECKING, Callable
import time
import threading
import contextlib
import uno
import unohelper
from com.sun.star.sheet import XRangeSelectionListener

from ooodev.globals.gbl_events import GblEvents
from ooodev.calc.calc_doc import CalcDoc
from ooodev.calc.calc_sheet_view import CalcSheetView
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.events.args.event_args import EventArgs
from ooodev.utils.helper.dot_dict import DotDict
from ooodev.utils import props as mProps
from ooodev.utils.data_type.range_obj import RangeObj
from ooodev.io.log.named_logger import NamedLogger

if TYPE_CHECKING:
    from com.sun.star.lang import EventObject
    from com.sun.star.sheet import RangeSelectionEvent


_SELECTION_MADE = None
_SELECTION_RESULT = None


class RangeSelector(EventsPartial):

    class _ExampleRangeListener(XRangeSelectionListener, EventsPartial, unohelper.Base):
        def __init__(
            self, view: CalcSheetView, auto_remove_listener: bool, single_cell_mode: bool, initial_value: str
        ):
            XRangeSelectionListener.__init__(self)
            EventsPartial.__init__(self)
            unohelper.Base.__init__(self)
            self._log = NamedLogger(name="RangeSelector._ExampleRangeListener")
            self._log.debug("RangeSelector._ExampleRangeListener.__init__")
            self.view = view
            self.auto_remove = auto_remove_listener
            self.single_cell_mode = single_cell_mode
            self.initial_value = initial_value
            self._removed = False

        def done(self, event: RangeSelectionEvent):
            if event.RangeDescriptor:
                dd = DotDict(
                    state="done",
                    result=event.RangeDescriptor,
                    event=event,
                    view=self.view,
                    rng_obj=None,
                    single_cell_mode=self.single_cell_mode,
                )
            else:
                dd = DotDict(
                    state="aborted",
                    result="",
                    event=event,
                    view=self.view,
                    rng_obj=None,
                    single_cell_mode=self.single_cell_mode,
                )
            dd.initial_value = self.initial_value
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
            self.trigger_event("AfterPopupRangeSelection", eargs)

        def aborted(self, event: RangeSelectionEvent):
            eargs = EventArgs(self)
            dd = DotDict(
                state="aborted",
                result="aborted",
                event=event,
                view=self.view,
                rng_obj=None,
                single_cell_mode=self.single_cell_mode,
                initial_value=self.initial_value,
            )
            eargs.event_data = dd
            if self.auto_remove and self._removed is False:
                self.view.remove_range_selection_listener(self)
                self._removed = True
            self.trigger_event("AfterPopupRangeSelection", eargs)

        def disposing(self, event: EventObject):
            pass

        def subscribe_range_select(self, cb: Callable[[Any, Any], None]) -> None:
            self.subscribe_event("AfterPopupRangeSelection", cb)

        def unsubscribe_range_select(self, cb: Callable[[Any, Any], None]) -> None:
            self.unsubscribe_event("AfterPopupRangeSelection", cb)

    def __init__(
        self,
        title: str = "Please select a range",
        close_on_mouse_release: bool = False,
        single_cell_mode: bool = False,
        initial_value: str = "",
    ):
        EventsPartial.__init__(self)
        self._gbl_events = GblEvents()
        self._log = NamedLogger(name="RangeSelection")
        self._log.debug("RangeSelector.__init__")
        self._title = title
        self._close_on_mouse_release = close_on_mouse_release
        self._single_cell_mode = single_cell_mode
        self._initial_value = initial_value
        self._init_cb()
        self._log.debug("RangeSelector.__init__() complete")

    def _init_cb(self) -> None:
        self._fn_on_range_sel = self._on_range_sel

    def _on_range_sel(self, src: Any, event: EventArgs):
        global _SELECTION_MADE, _SELECTION_RESULT

        _SELECTION_RESULT = event.event_data.rng_obj

        _SELECTION_MADE = True
        self._gbl_events.trigger_event("GlobalCalcRangeSelector", event)

    def get_range_selection(self, doc: CalcDoc) -> RangeObj | None:
        global _SELECTION_MADE, _SELECTION_RESULT
        self._log.debug("RangeSelector.get_range_selection() Entered")
        view = doc.get_view()
        self._log.debug("RangeSelector.get_range_selection() got view")
        ex_listener = RangeSelector._ExampleRangeListener(
            view=view,
            auto_remove_listener=True,
            single_cell_mode=self._single_cell_mode,
            initial_value=self._initial_value,
        )
        ex_listener.add_event_observers(self.event_observer)
        # ex_listener.add_event_observers(self._gbl_events.event_observer)
        self._log.debug("RangeSelector.get_range_selection() created listener")
        ex_listener.subscribe_range_select(self._fn_on_range_sel)
        self._log.debug("RangeSelector.get_range_selection() subscribed _fn_on_range_sel")
        _SELECTION_MADE = False
        _SELECTION_RESULT = None
        self._log.debug("RangeSelector.get_range_selection() set extra data")
        view.add_range_selection_listener(ex_listener)
        self._log.debug("RangeSelector.get_range_selection() added listener")
        if self._initial_value:
            props = mProps.Props.make_props(
                Title=self._title,
                CloseOnMouseRelease=self._close_on_mouse_release,
                SingleCellMode=self._single_cell_mode,
                InitialValue=self._initial_value,
            )
        else:
            props = mProps.Props.make_props(
                Title=self._title,
                CloseOnMouseRelease=self._close_on_mouse_release,
                SingleCellMode=self._single_cell_mode,
            )
        self._log.debug("RangeSelector.get_range_selection() made props")
        view.component.startRangeSelection(props)
        self._log.debug("RangeSelector.get_range_selection() started range selection")
        print("Make a selection in the document")
        tries = 0
        while not _SELECTION_MADE:
            tries += 1
            self._log.debug("RangeSelector.get_range_selection() waiting for selection")
            time.sleep(0.5)
            if tries > 120:
                self._log.warning("RangeSelector.get_range_selection() timeout")
                break  # break on 60 seconds max.
        result = _SELECTION_RESULT
        # ex_listener.remove_event_observer(self._gbl_events.event_observer)
        # ex_listener.remove_event_observer(self.event_observer)
        # del view.extra_data["selection_made"]
        # del view.extra_data["selection_result"]
        self._log.debug(f"RangeSelector.get_range_selection() results {result}")
        return result


class RangeSelectorThread(threading.Thread, EventsPartial):
    def __init__(
        self,
        title: str = "Please select a range",
        close_on_mouse_release: bool = False,
        single_cell_mode: bool = False,
        initial_value: str = "",
    ):
        threading.Thread.__init__(self)
        EventsPartial.__init__(self)
        self._stop_event = threading.Event()
        self._log = NamedLogger(name="RangeSelectorThread")
        self._rng_sel = RangeSelector(
            title=title,
            close_on_mouse_release=close_on_mouse_release,
            single_cell_mode=single_cell_mode,
            initial_value=initial_value,
        )
        self._rng_sel.add_event_observers(self.event_observer)
        self._result = None
        self._fn_on_sel_made = self._on_sel_made
        self._rng_sel.subscribe_event("AfterPopupRangeSelection", self._fn_on_sel_made)

    def _on_sel_made(self, src: Any, event: EventArgs):
        print("RangeSelectorThread._on_sel_made()")
        self._stop_event.set()

    def stop(self):
        self._stop_event.set()

    def stopped(self):
        return self._stop_event.is_set()

    def run(self):
        # doc = XSCRIPTCONTEXT.getDocument()
        # calc_doc = CalcDoc.get_doc_from_component(doc)
        try:
            if not self.stopped():
                calc_doc = CalcDoc.from_current_doc()
                self._result = self._rng_sel.get_range_selection(calc_doc)
        except Exception:
            self._log.error("Error in RangeSelectorThread.run()", exc_info=True)
            self._result = None
