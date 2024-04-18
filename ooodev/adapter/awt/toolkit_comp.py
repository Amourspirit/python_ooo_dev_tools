from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno
from com.sun.star.awt import XToolkit2

from ooodev.adapter.component_base import ComponentBase
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.adapter.awt.toolkit2_partial import Toolkit2Partial
from ooodev.adapter.awt.focus_events import FocusEvents
from ooodev.adapter.awt.top_window_events import TopWindowEvents
from ooodev.adapter.awt.key_handler_events import KeyHandlerEvents

if TYPE_CHECKING:
    from com.sun.star.awt import Toolkit
    from ooodev.loader.inst.lo_inst import LoInst


class ToolkitComp(ComponentBase, Toolkit2Partial, FocusEvents, TopWindowEvents, KeyHandlerEvents):
    """
    Class for managing PopupMenu Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XToolkit2) -> None:
        """
        Constructor

        Args:
            component (UnoControlDialog): UNO Component that supports `com.sun.star.awt.PopupMenu`` service.
        """
        # pylint: disable=no-member
        ComponentBase.__init__(self, component)
        Toolkit2Partial.__init__(self, component=component)
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        FocusEvents.__init__(self, trigger_args=generic_args, cb=self.__on_focus_add_remove_add_remove)
        TopWindowEvents.__init__(self, trigger_args=generic_args, cb=self.__on_top_window_add_remove_add_remove)
        KeyHandlerEvents.__init__(self, trigger_args=generic_args, cb=self.__on_keu_handler_add_remove_add_remove)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.awt.Toolkit",)

    # endregion Overrides

    # region Lazy Listeners

    def __on_focus_add_remove_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.component.addFocusListener(self.events_listener_focus)
        event.remove_callback = True

    def __on_top_window_add_remove_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.component.addTopWindowListener(self.events_listener_top_window)
        event.remove_callback = True

    def __on_keu_handler_add_remove_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.component.addKeyHandler(self.events_listener_key_handler)
        event.remove_callback = True

    # endregion Lazy Listeners

    @classmethod
    def from_lo(cls, lo_inst: LoInst | None = None) -> ToolkitComp:
        """
        Creates the instance from the Lo.

        Args:
            lo_inst (LoInst, optional): LoInst, Defaults to ``Lo.current_lo``.

        Returns:
            ToolkitComp: The instance.
        """
        # pylint: disable=import-outside-toplevel
        from ooodev.loader import lo as mLo

        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        inst = lo_inst.create_instance_mcf(XToolkit2, "com.sun.star.awt.Toolkit", raise_err=True)  # type: ignore
        return cls(inst)

    # region Properties

    @property
    def component(self) -> Toolkit:
        """Toolkit Component"""
        # pylint: disable=no-member
        return cast("Toolkit", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
