from __future__ import annotations
from typing import Any, Dict, List, Callable, TYPE_CHECKING
from ooodev.gui.menu.popup_menu import PopupMenu
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.loader import lo as mLo
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.events.args.event_args import EventArgs
from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.gui.menu.popup.menu_processor import MenuProcessor
from ooodev.gui.commands.cmd_info import CmdInfo

if TYPE_CHECKING:
    from ooodev.loader.inst.lo_inst import LoInst


class PopupCreator(LoInstPropsPartial, EventsPartial):
    """Class for creating popup menu"""

    def __init__(self, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            lo_inst (LoInst | None, optional): LibreOffice instance. Defaults to ``None``.
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        LoInstPropsPartial.__init__(self, lo_inst)
        EventsPartial.__init__(self)
        self._cmd_info = CmdInfo()

    def _insert_sub_menu(self, parent: PopupMenu, parent_menu_id: int, menus: list[dict[str, Any]]) -> None:
        """Insert submenu"""
        pm = PopupMenu.from_lo(lo_inst=self.lo_inst)
        eargs = EventArgs(self)
        eargs.event_data = {"popup_menu": pm}
        self.trigger_event("popup_created", eargs)

        for index, menu in enumerate(menus):
            submenu = menu.pop("submenu", False)
            mp = MenuProcessor(pm, self._cmd_info)
            mp.add_event_observers(self.event_observer)
            pop = mp.process(menu, index)
            if submenu:
                self._insert_sub_menu(pm, pop.menu_id, submenu)
        parent.set_popup_menu(parent_menu_id, pm)

    def create(self, menus: List[Dict[str, Any]]) -> PopupMenu:
        """
        Create popup menu.

        Args:
            menus (List[Dict[str, Any]]): Menu Data.
        """
        pm = PopupMenu.from_lo(lo_inst=self.lo_inst)
        eargs = EventArgs(self)
        eargs.event_data = {"popup_menu": pm}
        self.trigger_event("popup_created", eargs)
        for index, menu in enumerate(menus):
            submenu = menu.pop("submenu", False)
            mp = MenuProcessor(pm, self._cmd_info)
            mp.add_event_observers(self.event_observer)
            pop = mp.process(menu, index)
            if submenu:
                self._insert_sub_menu(pm, pop.menu_id, submenu)

        return pm

    def subscribe_popup_created(self, callback: Callable[[Any, EventArgs], None]) -> None:
        """
        Subscribe on popup created event.

        The callback ``event_data`` is a dictionary with keys:

        - ``popup_menu``: PopupMenu instance
        """
        self.subscribe_event("popup_created", callback)

    def unsubscribe_popup_created(self, callback: Callable[[Any, EventArgs], None]) -> None:
        """
        Unsubscribe on popup created event.
        """
        self.unsubscribe_event("popup_created", callback)

    def subscribe_before_process(self, callback: Callable[[Any, CancelEventArgs], None]) -> None:
        """
        Subscribe on before process event.

        The callback ``event_data`` is a dictionary with keys:

        - ``popup_menu``: PopupMenu instance
        - ``popup_item``: PopupItem instance
        """
        self.subscribe_event("before_process", callback)

    def subscribe_after_process(self, callback: Callable[[Any, EventArgs], None]) -> None:
        """
        Subscribe on after process event.

        The callback ``event_data`` is a dictionary with keys:

        - ``popup_menu``: PopupMenu instance
        - ``popup_item``: PopupItem instance
        """
        self.subscribe_event("after_process", callback)

    def unsubscribe_before_process(self, callback: Callable[[Any, CancelEventArgs], None]) -> None:
        """
        Unsubscribe on before process event.
        """
        self.unsubscribe_event("before_process", callback)

    def unsubscribe_after_process(self, callback: Callable[[Any, EventArgs], None]) -> None:
        """
        Unsubscribe on after process event.
        """
        self.unsubscribe_event("after_process", callback)
