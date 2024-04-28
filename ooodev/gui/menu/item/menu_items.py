from __future__ import annotations
from typing import cast, TYPE_CHECKING, Tuple
import uno
from com.sun.star.beans import PropertyValue
from ooodev.loader import lo as mLo
from ooodev.loader.inst.service import Service
from ooodev.adapter.container.index_access_comp import IndexAccessComp
from ooodev.utils import props as mProps
from ooodev.gui.menu.item.menu_item_sep import MenuItemSep
from ooodev.gui.menu.item.menu_item import MenuItem
from ooodev.gui.menu.item.menu_item_sub import MenuItemSub
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial


if TYPE_CHECKING:
    from ooodev.gui.menu.menu import Menu
    from com.sun.star.container import XIndexAccess
    from ooodev.loader.inst.lo_inst import LoInst


class MenuItems(LoInstPropsPartial, IndexAccessComp[Tuple[Tuple[PropertyValue, ...], ...]]):
    """Class for individual menu"""

    def __init__(self, component: XIndexAccess, menu: Menu, app: str | Service = "", lo_inst: LoInst | None = None):
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
        self._menu = menu

    def _get_name(self, name: str) -> str:
        return name[8:] if name.startswith(".custom:") else name

    def _get_by_cmd(self, name: str) -> Tuple[Tuple[PropertyValue, ...], ...] | None:
        """Get menu by command value"""
        name = self._get_name(name)
        menu = None
        for m in self:
            d = mProps.Props.data_to_dict(m)  # type: ignore
            cmd = d.get("CommandURL", "")
            if name == cmd:
                menu = m
                break
        return menu

    def __contains__(self, name: str) -> bool:
        """
        If exists name in menu.

        Args:
            name (str): Name of the menu item. If name starts with ``.custom:``, then  the ``.custom:`` part is removed.

        Returns:
            bool: ``True`` if exists, ``False`` otherwise.
        """

        name = self._get_name(name)
        exists = False
        for m in self:
            d = mProps.Props.data_to_dict(m)  # type: ignore
            cmd = d.get("CommandURL", "")
            if name == cmd:
                exists = True
                break
        return exists

    def __getitem__(self, index: int | str) -> MenuItemSep | MenuItem | MenuItemSub:
        """
        Index access.

        Args:
            index (int | str): Index of menu item or command value. if a string and starts with ``.custom:``, then  the ``.custom:`` part is removed.

        Raises:
            KeyError: If menu item not found.
            IndexError: If index out of range.

        Returns:
            MenuItem | MenuItemSep: Menu item or menu separator.
        """
        if isinstance(index, str):
            data = self._get_by_cmd(index)
            if data is None:
                raise KeyError(f"Menu item '{index}' not found")
        else:
            count = self.get_count()
            if index < 0 or index >= count:
                raise IndexError(f"Index out of range: {index}")
            data = super().get_by_index(index)
        sep = MenuItemSep(menu=self._menu, data=data, owner=self, app=self._app, lo_inst=self.lo_inst)
        if sep.menu_type == 1:
            return sep
        d = mProps.Props.data_to_dict(data)  # type: ignore
        itm_desc = cast("XIndexAccess", d.get("ItemDescriptorContainer", None))
        if itm_desc is None:
            count = 0
        else:
            count = itm_desc.getCount()
        if count == 0:
            return MenuItem(menu=self._menu, data=data, owner=self, app=self._app, lo_inst=self.lo_inst)
        return MenuItemSub(menu=self._menu, data=data, owner=self, app=self._app, lo_inst=self.lo_inst)
