from typing import Any
import pytest

# from ooodev.office.write import Write
if __name__ == "__main__":
    pytest.main([__file__])
# region    Sheet Methods


def test_highlight_range(loader) -> None:
    from ooodev.loader.lo import Lo
    from ooodev.office.calc import Calc
    from ooodev.utils.gui import GUI
    from ooodev.utils.color import CommonColor
    from ooodev.events.lo_events import event_ctx, EventArg
    from ooodev.events.calc_named_event import CalcNamedEvent
    from ooodev.events.args.calc.cell_args import CellArgs
    from ooodev.events.args.calc.cell_cancel_args import CellCancelArgs

    visible = False
    delay = 0  # 2_000

    is_adding_border = False
    is_highlighted = False

    def highlighting(source: Any, args: CellCancelArgs):
        args.event_data["color"] = CommonColor.GREEN
        args.event_data["headline"] = "Modified from callback"

    def highlighted(source: Any, args: CellArgs):
        nonlocal is_highlighted
        assert args.event_data["color"] == CommonColor.GREEN
        is_highlighted = True

    def adding_border(source: Any, args: CellCancelArgs):
        nonlocal is_adding_border
        is_adding_border = True

    assert loader is not None
    doc = Calc.create_doc(loader)
    try:
        rng_name = "B3:F8"
        headline = "Hello World!"
        with event_ctx(
            EventArg(CalcNamedEvent.CELLS_HIGH_LIGHTING, highlighting),
            EventArg(CalcNamedEvent.CELLS_HIGH_LIGHTED, highlighted),
            EventArg(CalcNamedEvent.CELLS_BORDER_ADDING, adding_border),
        ) as events:
            Lo.current_lo.add_event_observers(events)
            sheet = Calc.get_sheet(doc=doc, idx=0)
            rng = sheet.getCellRangeByName(rng_name)
            if visible:
                GUI.set_visible(visible=visible, doc=doc)
            first = Calc.highlight_range(sheet, headline, rng)
            Lo.delay(delay)
            assert first is not None
            result = Calc.get_string(cell=first)
            assert result == "Modified from callback"
            assert is_adding_border
            assert is_highlighted
            Lo.current_lo.remove_event_observer(events)

    finally:
        Lo.close_doc(doc)


def test_highlight_range_cancel(loader) -> None:
    from ooodev.loader.lo import Lo
    from ooodev.office.calc import Calc
    from ooodev.utils.gui import GUI
    from ooodev.events.lo_events import event_ctx, EventArg
    from ooodev.events.calc_named_event import CalcNamedEvent
    from ooodev.events.args.calc.cell_cancel_args import CellCancelArgs
    from ooodev.exceptions import ex as mEx

    visible = False
    delay = 0  # 2_000

    # raises and error due to callback setting cancel to True.

    def highlighting(source: Any, args: CellCancelArgs):
        args.cancel = True

    assert loader is not None
    doc = Calc.create_doc(loader)
    try:
        rng_name = "B3:F8"
        headline = "Hello World!"
        with event_ctx(EventArg(CalcNamedEvent.CELLS_HIGH_LIGHTING, highlighting)) as events:
            Lo.current_lo.add_event_observers(events)
            if visible:
                GUI.set_visible(visible=visible, doc=doc)
            sheet = Calc.get_sheet(doc=doc, idx=0)
            rng = sheet.getCellRangeByName(rng_name)
            with pytest.raises(mEx.CancelEventError):
                _ = Calc.highlight_range(sheet, headline, rng)

        # test outside of ctx manager
        first = Calc.highlight_range(sheet, headline, rng)
        result = Calc.get_string(cell=first)
        assert result == headline
        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)


def test_highlight_local_events(loader) -> None:
    from ooodev.loader.lo import Lo
    from ooodev.office.calc import Calc
    from ooodev.utils.gui import GUI
    from ooodev.utils.color import CommonColor
    from ooodev.events.lo_events import Events, LoEvents
    from ooodev.events.calc_named_event import CalcNamedEvent
    from ooodev.events.args.calc.cell_args import CellArgs
    from ooodev.events.args.calc.cell_cancel_args import CellCancelArgs

    visible = False
    delay = 0  # 2_000

    is_adding_border = False
    is_highlighted = False

    def highlighting(source: Any, args: CellCancelArgs):
        args.event_data["color"] = CommonColor.GREEN
        args.event_data["headline"] = "Modified from callback"

    def highlighted(source: Any, args: CellArgs):
        nonlocal is_highlighted
        assert args.event_data["color"] == CommonColor.GREEN
        is_highlighted = True

    def adding_border(source: Any, args: CellCancelArgs):
        # if args is None:
        #     return
        nonlocal is_adding_border
        is_adding_border = True

    assert loader is not None
    doc = Calc.create_doc(loader)
    try:
        if visible:
            GUI.set_visible(visible=visible, doc=doc)
        rng_name = "B3:F8"
        headline = "Hello World!"
        sheet = Calc.get_sheet(doc=doc, idx=0)
        rng = sheet.getCellRangeByName(rng_name)

        def do_work():
            # test event is a function call.
            # none of the events will be called outside of function.
            nonlocal sheet, headline, rng
            events = Events()
            events.on(CalcNamedEvent.CELLS_HIGH_LIGHTING, highlighting)
            events.on(CalcNamedEvent.CELLS_HIGH_LIGHTED, highlighted)
            events.on(CalcNamedEvent.CELLS_BORDER_ADDING, adding_border)
            Lo.current_lo.add_event_observers(events)
            return Calc.highlight_range(sheet, headline, rng)

        first = do_work()
        assert is_adding_border
        assert is_highlighted
        is_adding_border = False
        # test triggering event to make sure it has no effect outside of function.
        LoEvents().trigger(CalcNamedEvent.CELLS_BORDER_ADDING, None)
        assert is_adding_border is False
        Lo.delay(delay)
        assert first is not None
        result = Calc.get_string(cell=first)
        assert result == "Modified from callback"

    finally:
        Lo.close_doc(doc)


def test_highlight_events_destroy(loader) -> None:
    from ooodev.loader.lo import Lo
    from ooodev.office.calc import Calc
    from ooodev.utils.gui import GUI
    from ooodev.utils.color import CommonColor
    from ooodev.events.lo_events import Events, LoEvents
    from ooodev.events.calc_named_event import CalcNamedEvent
    from ooodev.events.args.calc.cell_args import CellArgs
    from ooodev.events.args.calc.cell_cancel_args import CellCancelArgs

    visible = False
    delay = 0  # 2_000

    is_adding_border = False
    is_highlighted = False

    def highlighting(source: Any, args: CellCancelArgs):
        args.event_data["color"] = CommonColor.GREEN
        args.event_data["headline"] = "Modified from callback"

    def highlighted(source: Any, args: CellArgs):
        nonlocal is_highlighted
        assert args.event_data["color"] == CommonColor.GREEN
        is_highlighted = True

    def adding_border(source: Any, args: CellCancelArgs):
        nonlocal is_adding_border
        is_adding_border = True

    assert loader is not None
    doc = Calc.create_doc(loader)
    try:
        if visible:
            GUI.set_visible(visible=visible, doc=doc)
        rng_name = "B3:F8"
        headline = "Hello World!"
        sheet = Calc.get_sheet(doc=doc, idx=0)
        rng = sheet.getCellRangeByName(rng_name)
        # by creating events and then destroying them it is
        # possible to register and then deregister callbacks
        events = Events()
        events.on(CalcNamedEvent.CELLS_HIGH_LIGHTING, highlighting)
        events.on(CalcNamedEvent.CELLS_HIGH_LIGHTED, highlighted)
        events.on(CalcNamedEvent.CELLS_BORDER_ADDING, adding_border)
        Lo.current_lo.add_event_observers(events)
        first = Calc.highlight_range(sheet, headline, rng)
        events = None
        _ = None
        # because events is now None the next call to LoEvents().trigger() will not
        # include this instance of events. This is accomplished by using weakref in LoEvents
        assert is_adding_border
        assert is_highlighted
        is_adding_border = False
        # test triggering event to make sure it has no effect outside of function.
        LoEvents().trigger(CalcNamedEvent.CELLS_BORDER_ADDING, None)
        assert is_adding_border is False
        Lo.delay(delay)
        assert first is not None
        result = Calc.get_string(cell=first)
        assert result == "Modified from callback"

    finally:
        Lo.close_doc(doc)
