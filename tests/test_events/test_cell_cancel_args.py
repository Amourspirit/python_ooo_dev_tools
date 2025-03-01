from typing import Any
import pytest
from hypothesis import given, settings
from hypothesis.strategies import text, characters, integers


@given(
    text(
        alphabet=characters(min_codepoint=95, max_codepoint=122, blacklist_characters=("`",)), min_size=3, max_size=25
    ),
    integers(min_value=-100, max_value=100),
)
@settings(max_examples=50)
def test_trigger_event(event_name_str: str, value: int):
    from ooodev.events.args.calc.cell_cancel_args import CellCancelArgs
    from ooodev.events.lo_events import Events

    cargs = CellCancelArgs(test_trigger_event)
    cargs.event_data = value
    cargs.sheet = None

    def triggered(source: Any, args: CellCancelArgs):
        args.cells = value
        assert source is test_trigger_event
        assert args.event_data == value
        assert args.source is test_trigger_event

    events = Events()
    events.on(event_name_str, triggered)
    events.trigger(event_name_str, cargs)
    assert cargs.event_name == event_name_str
    assert cargs.cancel is False
    assert cargs.cells == value
    assert cargs.sheet is None


@given(
    text(
        alphabet=characters(min_codepoint=95, max_codepoint=122, blacklist_characters=("`",)), min_size=3, max_size=25
    ),
    integers(min_value=-100, max_value=100),
)
@settings(max_examples=50)
def test_trigger_event_canceled(event_name_str: str, value: int):
    from ooodev.events.args.calc.cell_cancel_args import CellCancelArgs
    from ooodev.events.lo_events import Events

    cargs = CellCancelArgs(test_trigger_event_canceled)
    cargs.event_data = value
    cargs.sheet = None

    def triggered(source: Any, args: CellCancelArgs):
        args.cells = value
        assert source is test_trigger_event_canceled
        assert args.event_data == value
        assert args.source is test_trigger_event_canceled
        args.cancel = True
        assert args.cancel

    events = Events()
    events.on(event_name_str, triggered)
    events.trigger(event_name_str, cargs)
    assert cargs.event_name == event_name_str
    assert cargs.cancel
    assert cargs.cells == value
    assert cargs.sheet is None


@given(
    text(
        alphabet=characters(min_codepoint=95, max_codepoint=122, blacklist_characters=("`",)), min_size=3, max_size=25
    ),
    integers(min_value=-100, max_value=100),
)
@settings(max_examples=10)
def test_from_args(event_name_str: str, value: int):
    from ooodev.events.args.calc.cell_cancel_args import CellCancelArgs

    cargs = CellCancelArgs(test_from_args)
    cargs.event_data = value
    cargs._event_name = event_name_str
    cargs.sheet = None
    cargs.cells = value

    assert cargs.event_name == event_name_str
    assert cargs.event_data == value
    assert cargs.source is test_from_args
    assert cargs.cancel is False
    assert cargs.cells == value
    assert cargs.sheet is None

    e = CellCancelArgs.from_args(cargs)
    assert e.event_data == cargs.event_data
    assert e.source is cargs.source
    assert e.event_name == cargs.event_name
    assert e.cancel == cargs.cancel
    assert e.cells == cargs.cells
    assert e.sheet is cargs.sheet


@given(
    text(
        alphabet=characters(min_codepoint=95, max_codepoint=122, blacklist_characters=("`",)), min_size=3, max_size=25
    ),
    integers(min_value=-100, max_value=100),
)
@settings(max_examples=10)
def test_event_kv(key: str, value: Any) -> None:
    from ooodev.events.args.calc.cell_cancel_args import CellCancelArgs

    args = CellCancelArgs("random_event")
    args._event_name = "test_event_kv"

    with pytest.raises(KeyError):
        _ = args.get(key)

    assert args.get(key, None) is None
    assert args.set(key, value)
    assert args.has(key)
    val = args.get(key)
    assert val == value
    args.remove(key)
    assert args.has(key) == False
