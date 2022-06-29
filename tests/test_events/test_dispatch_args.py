from __future__ import annotations
from typing import Any

from hypothesis import given, settings
from hypothesis.strategies import text, characters, integers


@given(
    text(
        alphabet=characters(min_codepoint=95, max_codepoint=122, blacklist_characters=("`",)), min_size=3, max_size=25
    ),
    integers(min_value=-100, max_value=100),
)
@settings(max_examples=50)
def test_trigger_event(cmd: str, value: int):
    from ooodev.events.args.dispatch_args import DispatchArgs
    from ooodev.events.lo_events import Events

    dispathc_args = DispatchArgs(test_trigger_event, cmd)
    dispathc_args.event_data = value

    def triggered(source: Any, args: DispatchArgs):
        assert args.event_data == value
        assert args.source is test_trigger_event

    events = Events()
    events.on(cmd, triggered)
    events.trigger(cmd, dispathc_args)
    assert dispathc_args.event_name == cmd
    assert dispathc_args.cmd == cmd

@given(
    text(
        alphabet=characters(min_codepoint=95, max_codepoint=122, blacklist_characters=("`",)), min_size=3, max_size=25
    ),
    integers(min_value=-100, max_value=100),
)
@settings(max_examples=10)
def test_from_args(cmd:str, value: int):
    from ooodev.events.args.dispatch_args import DispatchArgs
    eargs = DispatchArgs(test_from_args, cmd)
    eargs.event_data = value
    eargs._event_name = cmd
    
    assert eargs.event_name == cmd
    assert eargs.event_data == value
    assert eargs.source is test_from_args

    e = DispatchArgs.from_args(eargs)
    assert e.event_data == eargs.event_data
    assert e.source is eargs.source
    assert e.event_name == eargs.event_name
    