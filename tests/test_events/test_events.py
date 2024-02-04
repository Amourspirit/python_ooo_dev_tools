from __future__ import annotations

import pytest
from typing import Any
from hypothesis import given, settings
from hypothesis.strategies import text, characters, integers

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.events.lo_events import Events, GenericArgs
from ooodev.events.args.event_args import EventArgs
from ooodev.events.args.cancel_event_args import CancelEventArgs


def test_event() -> None:
    fired = False

    def on_ev(source, event: EventArgs):
        nonlocal fired
        fired = True

    events = Events()
    events.on("ev", on_ev)
    events.trigger("ev", EventArgs("test"))
    assert fired is True


def test_event_cls_extra_arg() -> None:
    class MyClass:
        def __init__(self) -> None:
            self.events = Events(trigger_args=GenericArgs(self))
            self.fired = False
            self.events.on("ev", MyClass.on_ev)

        def trigger(self) -> None:
            self.events.trigger("ev", EventArgs(self.__class__.__name__))

        @staticmethod
        def on_ev(source, event: EventArgs, *args, **kwargs) -> None:
            instance = args[0]
            instance.fired = True

    clazz = MyClass()
    assert clazz.fired is False
    clazz.trigger()
    assert clazz.fired is True


def test_event_cls_source() -> None:
    class MyClass:
        def __init__(self) -> None:
            self.events = Events(source=self)
            self.fired = False
            self.events.on("ev", MyClass.on_ev)

        def trigger(self) -> None:
            self.events.trigger("ev", EventArgs(self.__class__.__name__))

        @staticmethod
        def on_ev(source, event: EventArgs) -> None:
            event.event_source.fired = True

    clazz = MyClass()
    assert clazz.fired is False
    clazz.trigger()
    assert clazz.fired is True


def test_event_cls_extra_kwarg() -> None:
    class MyClass:
        def __init__(self) -> None:
            self.events = Events(trigger_args=GenericArgs(myclass=self))
            self.fired = False
            self.events.on("ev", MyClass.on_ev)

        def trigger(self) -> None:
            self.events.trigger(event_name="ev", event_args=EventArgs(self.__class__.__name__))

        @staticmethod
        def on_ev(source, event: EventArgs, *args, **kwargs) -> None:
            instance = kwargs["myclass"]
            instance.fired = True

    clazz = MyClass()
    assert clazz.fired is False
    clazz.trigger()
    assert clazz.fired is True


def test_event_cls() -> None:
    class MyClass:
        def __init__(self) -> None:
            self.events = Events()
            self.fired = False
            self.events.on("ev", MyClass.on_ev)

        def trigger(self) -> None:
            eargs = EventArgs(self.__class__.__name__)
            eargs.event_data = self
            self.events.trigger("ev", eargs)

        @staticmethod
        def on_ev(source, event: EventArgs) -> None:
            event.event_data.fired = True

    clazz = MyClass()
    assert clazz.fired is False
    clazz.trigger()
    assert clazz.fired is True


def test_event_new_doc(loader) -> None:
    from ooodev.loader.lo import Lo
    from ooodev.office.write import Write
    from ooodev.events.write_named_event import WriteNamedEvent

    fired_creating = False
    fired_created = False

    def on_write_creating(source: Any, e: CancelEventArgs) -> None:
        nonlocal fired_creating
        fired_creating = True

    def on_write_created(source: Any, e: EventArgs) -> None:
        nonlocal fired_created
        fired_created = True

    events = Events()
    events.on(WriteNamedEvent.DOC_CREATING, on_write_creating)
    events.on(WriteNamedEvent.DOC_CREATED, on_write_created)
    Lo.current_lo.add_event_observers(events)
    _ = Write.create_doc(loader)
    assert fired_creating is True
    assert fired_created is True


@given(
    text(
        alphabet=characters(min_codepoint=95, max_codepoint=122, blacklist_characters=("`",)), min_size=3, max_size=25
    ),
    integers(min_value=-100, max_value=100),
)
@settings(max_examples=10)
def test_event_kv(key: str, value: Any) -> None:
    args = EventArgs("random_event")
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
