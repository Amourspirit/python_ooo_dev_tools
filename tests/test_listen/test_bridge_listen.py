from __future__ import annotations
import pytest
import types
from typing import Any, TYPE_CHECKING
from ooodev.events.args.event_args import EventArgs
from ooodev.events.event_singleton import _Events

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.listeners.x_event_adapter import XEventAdapter
from ooodev.adapter.lang.event_events import EventEvents

if TYPE_CHECKING:
    # only need types in design time and not at run time.
    from com.sun.star.lang import EventObject


def test_bridge(loader) -> None:
    # does not assert anything.
    # just ensures adding a bridge event listener does not throw errors.
    from ooodev.loader import Lo

    def disposing_bridge(src: XEventAdapter, event: EventObject) -> None:
        print("Test Bridge - Office bridge has gone!!")
        return

    bridge_listen = XEventAdapter()
    bridge_listen.disposing = types.MethodType(disposing_bridge, bridge_listen)
    Lo.bridge.addEventListener(bridge_listen)
    assert True


def test_bridge_events(loader) -> None:
    # does not assert anything.
    # just ensures adding a bridge event listener does not throw errors.
    from ooodev.loader import Lo

    # import must be outside of function in order to work here.
    # from ooodev.adapter.lang.event_events import EventEvents

    def disposing_bridge(src: Any, event: EventArgs, *args: Any, **kwargs: Any) -> None:
        print("Test Bridge Events - Office bridge has gone!!")
        return

    bridge_listen = EventEvents(subscriber=Lo.bridge)
    bridge_listen.add_event_disposing(disposing_bridge)

    # this works also
    # Lo.bridge.addEventListener(bridge_listen.events_listener_event)

    assert True


def test_bridge_global_events(loader) -> None:
    # does not assert anything.
    # just ensures adding a bridge event listener does not throw errors.
    from ooodev.events.lo_named_event import LoNamedEvent

    # import must be outside of function in order to work here.
    # from ooodev.adapter.lang.event_events import EventEvents

    def disposing_bridge(src: Any, event: EventArgs, *args: Any, **kwargs: Any) -> None:
        print("Test Bridge Global Events - Office bridge has gone!!")
        return

    _Events().on(LoNamedEvent.BRIDGE_DISPOSED, disposing_bridge)

    assert True
