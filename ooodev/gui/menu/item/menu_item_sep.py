from __future__ import annotations
from ooodev.gui.menu.item.menu_item_base import MenuItemBase
from ooodev.gui.menu.item.menu_item_kind import MenuItemKind


class MenuItemSep(MenuItemBase):
    """Menu Item Separator"""

    def __repr__(self) -> str:
        return f"<MenuItemSep(kind={str(self.item_kind)})>"

    @property
    def item_kind(self) -> MenuItemKind:
        """
        Get item kind.

        Returns:
            MenuItemKind: ``MenuItemKind.SEP``.

        Hint:
            - ``MenuItemKind`` is an enum and can be imported from ``ooodev.gui.menu.item.menu_item_kind``.
        """
        return MenuItemKind.SEP
