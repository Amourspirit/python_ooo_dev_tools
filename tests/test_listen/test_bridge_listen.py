from __future__ import annotations
import pytest
import types
from typing import TYPE_CHECKING

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.listeners.x_event_adapter import XEventAdapter

if TYPE_CHECKING:
    # only need types in design time and not at run time.
    from com.sun.star.lang import EventObject

def test_bridge(loader) -> None:
    # does not assert anything.
    # just ensures adding a bridge event listeren does not throw errors.
    from ooodev.utils.lo import Lo
    def disposing_bridge(src: XEventAdapter, event: EventObject) -> None:
        print("Office bridge has gone!!")
        return
    bridge_listen = XEventAdapter()
    bridge_listen.disposing = types.MethodType(disposing_bridge, bridge_listen)
    Lo.bridge.addEventListener(bridge_listen)
    assert True
