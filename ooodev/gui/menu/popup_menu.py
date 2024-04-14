from __future__ import annotations
from re import sub
import uno
from com.sun.star.awt import XPopupMenu
from ooo.dyn.awt.menu_item_type import MenuItemType

from ooodev.adapter.awt.popup_menu_comp import PopupMenuComp


class PopupMenu(PopupMenuComp):

    def __init__(self, component: XPopupMenu) -> None:
        super().__init__(component)

    # region Find methods
    def get_max_menu_id(self) -> int:
        """
        Gets the maximum menu id.

        Returns:
            int: The maximum menu id.
        """
        max_id = -1
        for i in self:
            if i > max_id:
                max_id = i
        return max_id

    def find_item_pos(self, cmd: str, search_sub_menu: bool = False) -> int:
        """
        Find item position by command.

        Args:
            cmd (str): A menu command such as ``.uno:Copy``.
            search_sub_menu (bool, optional): Search in sub menus. Defaults to ``False``.

        Returns:
            int: The position of the menu item. If not found, return -1.

        See Also:
            - :meth:`find_item_menu_id`
        """

        def search(pop_mnu: PopupMenu, str_cmd: str) -> int:
            nonlocal search_sub_menu
            result = -1
            cmd = str_cmd.casefold()
            for i, menu_id in enumerate(pop_mnu):
                if search_sub_menu:
                    submenu = pop_mnu.get_popup_menu(menu_id)
                    if submenu is not None:
                        result = search(submenu, cmd)
                        if result != -1:
                            break
                menu_type = pop_mnu.get_item_type(i)
                if menu_type == MenuItemType.SEPARATOR:
                    continue
                command = pop_mnu.get_command(menu_id)
                if cmd and cmd == command.casefold():
                    result = i
                    break
            return result

        result = -1
        if cmd.startswith(".custom:"):
            cmd = cmd[8:]
        if not cmd:
            return result
        result = search(self, cmd)
        return result

    def find_item_menu_id(self, cmd: str, search_sub_menu: bool = False) -> int:
        """
        Find item menu id by command.

        Args:
            cmd (str): A menu command such as ``.uno:Copy``.
            search_sub_menu (bool, optional): Search in sub menus. Defaults to ``False``.

        Returns:
            int: The id of the menu item. If not found, return -1.

        See Also:
            - :meth:`find_item_pos`
        """
        result = self.find_item_pos(cmd, search_sub_menu)
        if result == -1:
            return -1
        return self.get_item_id(result)

    # endregion Find methods
