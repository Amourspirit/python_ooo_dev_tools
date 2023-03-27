from __future__ import annotations
from typing import Any
import pytest
from hypothesis import given, example, settings
from hypothesis.strategies import text, characters, integers


@given(text(max_size=25))
@example("")
@settings(max_examples=25)
def test_event_name(test_str: str):
    from ooodev.events.args.cancel_event_args import CancelEventArgs

    e = CancelEventArgs(test_event_name)
    e._event_name = test_str
    assert e.event_name == test_str


def test_source():
    from ooodev.events.args.cancel_event_args import CancelEventArgs

    e = CancelEventArgs(test_source)
    assert e.source is test_source


@given(
    text(
        alphabet=characters(min_codepoint=95, max_codepoint=122, blacklist_characters=("`",)), min_size=3, max_size=25
    ),
    integers(min_value=-100, max_value=100),
)
@settings(max_examples=50)
def test_trigger_event(event_name_str: str, value: int):
    from ooodev.events.args.cancel_event_args import CancelEventArgs
    from ooodev.events.lo_events import Events

    cargs = CancelEventArgs(test_trigger_event)
    cargs.event_data = value

    def triggered(source: Any, args: CancelEventArgs):
        assert args.event_data == value
        assert args.source is test_trigger_event

    events = Events()
    events.on(event_name_str, triggered)
    events.trigger(event_name_str, cargs)
    assert cargs.event_name == event_name_str
    assert cargs.cancel is False


@given(
    text(
        alphabet=characters(min_codepoint=95, max_codepoint=122, blacklist_characters=("`",)), min_size=3, max_size=25
    ),
    integers(min_value=-100, max_value=100),
)
@settings(max_examples=50)
def test_trigger_event_canceled(event_name_str: str, value: int):
    from ooodev.events.args.cancel_event_args import CancelEventArgs
    from ooodev.events.lo_events import Events

    cargs = CancelEventArgs(test_trigger_event_canceled)
    cargs.event_data = value

    def triggered(source: Any, args: CancelEventArgs):
        assert args.event_data == value
        assert args.source is test_trigger_event_canceled
        args.cancel = True
        assert args.cancel

    events = Events()
    events.on(event_name_str, triggered)
    events.trigger(event_name_str, cargs)
    assert cargs.event_name == event_name_str
    assert cargs.cancel


@given(
    text(
        alphabet=characters(min_codepoint=95, max_codepoint=122, blacklist_characters=("`",)), min_size=3, max_size=25
    ),
    integers(min_value=-100, max_value=100),
)
@settings(max_examples=10)
def test_from_args(event_name_str: str, value: int):
    from ooodev.events.args.cancel_event_args import CancelEventArgs

    cargs = CancelEventArgs(test_from_args.__qualname__)
    cargs.event_data = value
    cargs._event_name = event_name_str

    assert cargs.event_name == event_name_str
    assert cargs.event_data == value
    assert cargs.source is test_from_args.__qualname__
    assert cargs.cancel is False

    e = CancelEventArgs.from_args(cargs)
    assert e.event_data == cargs.event_data
    assert e.source is cargs.source
    assert e.event_name == cargs.event_name
    assert e.cancel == cargs.cancel


@given(
    text(
        alphabet=characters(min_codepoint=95, max_codepoint=122, blacklist_characters=("`",)), min_size=3, max_size=25
    ),
    integers(min_value=-100, max_value=100),
)
@settings(max_examples=10)
def test_event_kv(key: str, value: Any) -> None:
    from ooodev.events.args.cancel_event_args import CancelEventArgs

    cargs = CancelEventArgs("random_event")
    cargs._event_name = "test_event_kv"

    with pytest.raises(KeyError):
        _ = cargs.get(key)

    assert cargs.get(key, None) is None
    assert cargs.set(key, value)
    assert cargs.has(key)
    val = cargs.get(key)
    assert val == value
    cargs.remove(key)
    assert cargs.has(key) == False
