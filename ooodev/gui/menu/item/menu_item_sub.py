from __future__ import annotations
from typing import Tuple, TYPE_CHECKING
import uno
from com.sun.star.beans import PropertyValue

from ooodev.adapter.container.index_access_comp import IndexAccessComp
from ooodev.gui.menu.item.menu_item import MenuItem
from ooodev.gui.menu.item.menu_item_kind import MenuItemKind
from ooodev.loader.inst.service import Service

if TYPE_CHECKING:
    from ooodev.gui.menu.menu import Menu
    from ooodev.loader.inst.lo_inst import LoInst


class MenuItemSub(MenuItem):
    """Menu Item"""

    def __init__(
        self,
        *,
        menu: Menu,
        data: Tuple[Tuple[PropertyValue, ...], ...],
        owner: IndexAccessComp,
        app: str | Service = "",
        lo_inst: LoInst | None = None,
    ):
        """
        Constructor

        Args:
            data (Tuple[Tuple[PropertyValue, ...], ...]): UNO Object containing menu item properties.
            owner (IndexAccessComp): Parent menu.
            app (str | Service, optional): Name LibreOffice module. Defaults to "".
            lo_inst (LoInst | None, optional): Lo Instance. Defaults to Current Lo Instance.
        """
        super().__init__(data=data, menu=menu, owner=owner, app=app, lo_inst=lo_inst)
        self._sub_menu = None

    @property
    def sub_menu(self) -> Menu:
        """Get sub menu"""
        if self._sub_menu is None:
            self._sub_menu = self.menu[self.command]
        return self._sub_menu

    @property
    def item_kind(self) -> MenuItemKind:
        """
        Get item kind.

        Returns:
            MenuItemKind: ``MenuItemKind.ITEM_SUBMENU``.

        Hint:
            - ``MenuItemKind`` is an enum and can be imported from ``ooodev.gui.menu.item.menu_item_kind``.
        """
        return MenuItemKind.ITEM_SUBMENU
