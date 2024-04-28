from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
from ooodev.adapter.component_base import ComponentBase
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.adapter.frame.desktop2_partial import Desktop2Partial
from ooodev.adapter.frame.terminate_events import TerminateEvents
from ooodev.adapter.frame.frame_action_events import FrameActionEvents
from ooodev.adapter.frame.dispatch_provider_interception_partial import DispatchProviderInterceptionPartial

if TYPE_CHECKING:
    from com.sun.star.frame import theDesktop  # singleton
    from ooodev.loader.inst.lo_inst import LoInst


class TheDesktopComp(
    ComponentBase, Desktop2Partial, DispatchProviderInterceptionPartial, TerminateEvents, FrameActionEvents
):
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
        DispatchProviderInterceptionPartial.__init__(self, component=component, interface=None)
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
    @classmethod
    def from_lo(cls, lo_inst: LoInst | None = None) -> TheDesktopComp:
        """
        Get the singleton instance from the Lo.

        Args:
            lo_inst (LoInst, optional): LoInst, Defaults to ``Lo.current_lo``.

        Returns:
            TheDesktopComp: The instance.
        """
        # pylint: disable=import-outside-toplevel
        from ooodev.loader import lo as mLo

        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        key = "com.sun.star.frame.theDesktop"
        if key in lo_inst.cache:
            return cast(TheDesktopComp, lo_inst.cache[key])
        factory = lo_inst.get_singleton("/singletons/com.sun.star.frame.theDesktop")  # type: ignore
        if factory is None:
            raise ValueError("Could not get theDesktop singleton.")
        inst = cls(factory)
        lo_inst.cache[key] = inst
        return cast(TheDesktopComp, inst)

    # region Properties
    @property
    def component(self) -> theDesktop:
        """theDesktop Component"""
        # pylint: disable=no-member
        return cast("theDesktop", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
