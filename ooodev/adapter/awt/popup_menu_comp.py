from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno
from com.sun.star.awt import XPopupMenu

from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.awt.popup_menu_partial import PopupMenuPartial
from ooodev.adapter.awt.menu_events import MenuEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs

if TYPE_CHECKING:
    from com.sun.star.awt import PopupMenu
    from ooodev.loader.inst.lo_inst import LoInst


class PopupMenuComp(ComponentBase, PopupMenuPartial, MenuEvents):
    """
    Class for managing PopupMenu Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XPopupMenu) -> None:
        """
        Constructor

        Args:
            component (UnoControlDialog): UNO Component that supports `com.sun.star.awt.PopupMenu`` service.
        """
        # pylint: disable=no-member
        ComponentBase.__init__(self, component)
        PopupMenuPartial.__init__(self, component=component)
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        MenuEvents.__init__(self, trigger_args=generic_args, cb=self.__on_menu_add_remove_add_remove)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.awt.PopupMenu",)

    # endregion Overrides

    # region Lazy Listeners

    def __on_menu_add_remove_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.component.addMenuListener(self.events_listener_menu)
        event.remove_callback = True

    # endregion Lazy Listeners

    @classmethod
    def from_lo(cls, lo_inst: LoInst | None = None) -> PopupMenuComp:
        """
        Creates the instance from the Lo.

        Args:
            lo_inst (LoInst, optional): LoInst, Defaults to ``Lo.current_lo``.

        Returns:
            PopupMenuComp: The instance.
        """
        # pylint: disable=import-outside-toplevel
        from ooodev.loader import lo as mLo

        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        inst = lo_inst.create_instance_mcf(XPopupMenu, "com.sun.star.awt.PopupMenu", raise_err=True)  # type: ignore
        return cls(inst)

    # region Properties

    @property
    def component(self) -> PopupMenu:
        """PopupMenu Component"""
        # pylint: disable=no-member
        return cast("PopupMenu", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
