from __future__ import annotations

import pytest
import types
from typing import Any
import sys

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.events.lo_events import Events
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
            self.events = Events()
            self.fired = False
            self.events.on("ev", MyClass.on_ev)

        def trigger(self) -> None:
            self.events.trigger("ev", EventArgs(self.__class__.__name__), self)

        @staticmethod
        def on_ev(source, event: EventArgs, instance: MyClass) -> None:
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
            self.events = Events()
            self.fired = False
            self.events.on("ev", MyClass.on_ev)

        def trigger(self) -> None:
            self.events.trigger(event_name="ev", event_args=EventArgs(self.__class__.__name__), instance=self)

        @staticmethod
        def on_ev(source, event: EventArgs, instance: MyClass) -> None:
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
    from ooodev.utils.lo import Lo
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
    doc = Write.create_doc(loader)
    assert fired_creating is True
    assert fired_created is True
