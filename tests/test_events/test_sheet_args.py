from __future__ import annotations
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
def test_trigger_event(name: str, value: int):
    from ooodev.events.args.calc.sheet_args import SheetArgs
    from ooodev.events.lo_events import Events

    eargs = SheetArgs(test_trigger_event)
    eargs.sheet = None
    eargs.doc = None

    def triggered(source: Any, args: SheetArgs):
        assert source is test_trigger_event
        eargs.event_data = value
        eargs.index = value
        eargs.name = name
        assert args.event_data == value
        assert args.source is test_trigger_event

    events = Events()
    events.on(name, triggered)
    events.trigger(name, eargs)
    assert eargs.event_name == name
    assert eargs.event_data == value
    assert eargs.sheet is None
    assert eargs.doc is None
    assert eargs.index == value
    assert eargs.name == name


@given(
    text(
        alphabet=characters(min_codepoint=95, max_codepoint=122, blacklist_characters=("`",)), min_size=3, max_size=25
    ),
    integers(min_value=-100, max_value=100),
)
@settings(max_examples=10)
def test_from_args(name: str, value: int):
    from ooodev.events.args.calc.sheet_args import SheetArgs

    eargs = SheetArgs(test_from_args)
    eargs.event_data = value
    eargs._event_name = name
    eargs.sheet = None
    eargs.index = value
    eargs.name = name
    eargs.doc = None

    assert eargs.event_name == name
    assert eargs.event_data == value
    assert eargs.source is test_from_args
    assert eargs.doc is None
    assert eargs.index == value
    assert eargs.name == name

    e = SheetArgs.from_args(eargs)
    assert e.event_data == eargs.event_data
    assert e.source is eargs.source
    assert e.event_name == eargs.event_name
    assert e.sheet is eargs.sheet
    assert e.doc is eargs.doc
    assert e.index is eargs.index
    assert e.name is eargs.name


@given(
    text(
        alphabet=characters(min_codepoint=95, max_codepoint=122, blacklist_characters=("`",)), min_size=3, max_size=25
    ),
    integers(min_value=-100, max_value=100),
)
@settings(max_examples=10)
def test_event_kv(key: str, value: Any) -> None:
    from ooodev.events.args.calc.sheet_args import SheetArgs

    args = SheetArgs("random_event")
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
