from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
from ooodev.adapter.component_base import ComponentBase
from ooodev.events.args.listener_event_args import ListenerEventArgs

from ooodev.adapter.frame.desktop2_partial import Desktop2Partial
from ooodev.adapter.frame.terminate_events import TerminateEvents
from ooodev.adapter.frame.frame_action_events import FrameActionEvents

if TYPE_CHECKING:
    from com.sun.star.frame import theDesktop  # singleton


class TheDesktopComp(ComponentBase, Desktop2Partial, TerminateEvents, FrameActionEvents):
    """
    Class for managing theDesktop Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: theDesktop) -> None:
        """
        Constructor

        Args:
            component (theDesktop): UNO Component that implements ``com.sun.star.frame.theDesktop`` service.
        """
        ComponentBase.__init__(self, component)
        Desktop2Partial.__init__(self, component=component, interface=None)
        # pylint: disable=no-member
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        TerminateEvents.__init__(self, trigger_args=generic_args, cb=self._on_key_terminate_events_add_remove)
        FrameActionEvents.__init__(self, trigger_args=generic_args, cb=self._on_frame_action_events_add_remove)

    # region Lazy Listeners
    def _on_key_terminate_events_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.component.addTerminateListener(self.events_listener_terminate)
        event.remove_callback = True

    def _on_frame_action_events_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.component.addFrameActionListener(self.events_listener_frame_action)
        event.remove_callback = True

    # endregion Lazy Listeners

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ()

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> theDesktop:
        """theDesktop Component"""
        # pylint: disable=no-member
        return cast("theDesktop", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
