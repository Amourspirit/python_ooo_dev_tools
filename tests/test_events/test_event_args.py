from __future__ import annotations
from typing import Any

from hypothesis import given, example, settings
from hypothesis.strategies import text, characters, integers

from ooodev.events.args.event_args import EventArgs

@given(text(max_size=25))
@example("")
@settings(max_examples=25)
def test_event_name(test_str: str):
    e = EventArgs(test_event_name)
    e._event_name = test_str
    assert e.event_name == test_str


def test_source():
    e = EventArgs(test_source)
    assert e.source is test_source

@given(
    text(
        alphabet=characters(min_codepoint=95, max_codepoint=122, blacklist_characters=("`",)), min_size=3, max_size=25
    ),
    integers(min_value=-100, max_value=100),
)
@settings(max_examples=50)
def test_trigger_event(event_name_str: str, value: int):
    from ooodev.events.lo_events import Events

    eargs = EventArgs(test_trigger_event)
    eargs.event_data = value

    def triggered(source: Any, args: EventArgs):
        assert args.event_data == value
        assert args.source is test_trigger_event

    events = Events()
    events.on(event_name_str, triggered)
    events.trigger(event_name_str, eargs)
    assert eargs.event_name == event_name_str
