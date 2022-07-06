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
    from ooodev.events.args.dispatch_cancel_args import DispatchCancelArgs
    from ooodev.events.lo_events import Events

    dispatch_cancel_args = DispatchCancelArgs(test_trigger_event, cmd)
    dispatch_cancel_args.event_data = value

    def triggered(source: Any, args: DispatchCancelArgs):
        args.cancel = True
        assert args.event_data == value
        assert args.source is test_trigger_event

    events = Events()
    events.on(cmd, triggered)
    events.trigger(cmd, dispatch_cancel_args)
    assert dispatch_cancel_args.event_name == cmd
    assert dispatch_cancel_args.cmd == cmd
    assert dispatch_cancel_args.cancel is True

@given(
    text(
        alphabet=characters(min_codepoint=95, max_codepoint=122, blacklist_characters=("`",)), min_size=3, max_size=25
    ),
    integers(min_value=-100, max_value=100),
)
@settings(max_examples=10)
def test_from_args(cmd:str, value: int):
    from ooodev.events.args.dispatch_cancel_args import DispatchCancelArgs
    dispatch_cancel_args = DispatchCancelArgs(test_from_args, cmd)
    dispatch_cancel_args.event_data = value
    dispatch_cancel_args._event_name = cmd
    dispatch_cancel_args.cmd = cmd
    
    assert dispatch_cancel_args.event_name == cmd
    assert dispatch_cancel_args.event_data == value
    assert dispatch_cancel_args.source is test_from_args
    assert dispatch_cancel_args.cancel is False
    assert dispatch_cancel_args.cmd == cmd

    e = DispatchCancelArgs.from_args(dispatch_cancel_args)
    assert e.event_data == dispatch_cancel_args.event_data
    assert e.source is dispatch_cancel_args.source
    assert e.event_name == dispatch_cancel_args.event_name
    assert e.cancel == dispatch_cancel_args.cancel
    assert e.cmd == dispatch_cancel_args.cmd
    