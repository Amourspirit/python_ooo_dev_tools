from __future__ import annotations
from typing import TYPE_CHECKING, Tuple
import uno
from com.sun.star.beans import PropertyValue
from ooodev.loader import lo as mLo
from ooodev.loader.inst.service import Service
from ooodev.adapter.container.index_access_comp import IndexAccessComp
from ooodev.utils import props as mProps
from ooodev.gui.menu.item.menu_item_sep import MenuItemSep
from ooodev.gui.menu.item.menu_item import MenuItem
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial


if TYPE_CHECKING:
    from com.sun.star.container import XIndexAccess
    from ooodev.loader.inst.lo_inst import LoInst


class MenuItems(LoInstPropsPartial, IndexAccessComp[Tuple[Tuple[PropertyValue, ...], ...]]):
    """Class for individual menu"""

    def __init__(self, component: XIndexAccess, app: str | Service = "", lo_inst: LoInst | None = None):
        """
        Constructor

        Args:
            component (XIndexAccess): UNO Object.
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        LoInstPropsPartial.__init__(self, lo_inst)
        IndexAccessComp.__init__(self, component)
        self._app = str(app)

    def _get_by_cmd(self, name) -> Tuple[Tuple[PropertyValue, ...], ...] | None:
        """Get menu by command value"""
        menu = None
        for m in self:
            d = mProps.Props.data_to_dict(m)  # type: ignore
            cmd = d.get("CommandURL", "")
            if name == cmd:
                menu = m
                break
        return menu

    def __contains__(self, name) -> bool:
        """If exists name in menu"""
        exists = False
        for m in self:
            d = mProps.Props.data_to_dict(m)  # type: ignore
            cmd = d.get("CommandURL", "")
            if name == cmd:
                exists = True
                break
        return exists

    def __getitem__(self, index: int | str):
        """Index access"""
        if isinstance(index, str):
            data = self._get_by_cmd(index)
            if data is None:
                raise KeyError(f"Menu item '{index}' not found")
        else:
            data = super().get_by_index(index)
        sep = MenuItemSep(data=data, owner=self, app=self._app, lo_inst=self.lo_inst)
        if sep.menu_type == 1:
            return sep
        return MenuItem(data=data, owner=self, app=self._app, lo_inst=self.lo_inst)
