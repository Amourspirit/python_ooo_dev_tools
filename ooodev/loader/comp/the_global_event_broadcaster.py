from __future__ import annotations
from typing import Any
from ooodev.adapter.frame.the_global_event_broadcaster_comp import TheGlobalEventBroadcasterComp


class TheGlobalEventBroadcaster(TheGlobalEventBroadcasterComp):
    """
    Class for managing theGlobalEventBroadcaster singleton Class.
    """

    def __init__(self, component: Any) -> None:
        """
        Constructor

        Args:
            component (Any): UNO Component that implements ``com.sun.star.frame.theGlobalEventBroadcaster`` service.
        """
        TheGlobalEventBroadcasterComp.__init__(self, component=component)
