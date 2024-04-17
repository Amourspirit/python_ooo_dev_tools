from __future__ import annotations
from typing import TYPE_CHECKING, Tuple
import uno
from com.sun.star.awt import XPopupMenu
from ooo.dyn.awt.menu_item_type import MenuItemType

from logging import DEBUG
from ooodev.adapter.component_prop import ComponentProp
from ooodev.io.log.named_logger import NamedLogger
from ooodev.io.log import logging as logger
from ooodev.adapter.awt.popup_menu_comp import PopupMenuComp
from ooodev.utils.lru_cache import LRUCache
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.macro.script.macro_script import MacroScript
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.kind.menu_item_style_kind import MenuItemStyleKind
    from ooodev.loader.inst.lo_inst import LoInst


class PopupMenu(LoInstPropsPartial, PopupMenuComp):
    """Popup Menu Class."""

    def __init__(self, component: XPopupMenu, lo_inst: LoInst | None = None) -> None:
        """
        Initializes the instance.

        Args:
            component (XPopupMenu): The popup menu component.
            lo_inst (LoInst, optional): LoInst, Defaults to ``Lo.current_lo``.
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        LoInstPropsPartial.__init__(self, lo_inst)
        PopupMenuComp.__init__(self, component)
        log_level = logger.get_log_level()
        if log_level == DEBUG:
            self._logger = NamedLogger(f"{self.__class__.__name__} - {id(self)}")
        else:
            self._logger = NamedLogger(self.__class__.__name__)

        self._cache = LRUCache(30)

    # region Find methods
    def get_max_menu_id(self) -> int:
        """
        Gets the maximum menu id.

        Returns:
            int: The maximum menu id.
        """
        key = "get_max_menu_id"
        if key in self._cache:
            return self._cache[key]
        max_id = -1
        for i in self:
            if i > max_id:
                max_id = i
        self._cache[key] = max_id
        return max_id

    def find_item_pos(self, cmd: str, search_sub_menu: bool = False) -> Tuple[int, PopupMenu | None]:
        """
        Find item position by command.

        Args:
            cmd (str): A menu command such as ``.uno:Copy``.
            search_sub_menu (bool, optional): Search in sub menus. Defaults to ``False``.

        Returns:
            Tuple[int, PopupMenu | None]: The position of the menu item and the Popup Menu that command was found in. If not found, return ``(-1, None)``.

        See Also:
            - :meth:`find_item_menu_id`
        """
        key = f"find_item_pos_{cmd}_{search_sub_menu}"
        if key in self._cache:
            return self._cache[key]

        def search(pop_mnu: PopupMenu, str_cmd: str) -> Tuple[int, PopupMenu | None]:
            nonlocal search_sub_menu
            result = -1
            submenu = None
            if self._logger.is_debug:
                self._logger.debug(f"Searching for command: {str_cmd}")
            cmd = str_cmd.casefold()
            for i, menu_id in enumerate(pop_mnu):
                if search_sub_menu:
                    if self._logger.is_debug:
                        self._logger.debug(f"Searching {str_cmd} in sub menu: {menu_id}")
                    submenu = pop_mnu.get_popup_menu(menu_id)
                    if submenu is not None and len(submenu) > 0:
                        # turn of caching for sub menu when searching. No real value.
                        submenu.cache.capacity = 0
                        if self._logger.is_debug:
                            self._logger.debug(f"Found a sub Menu {str_cmd}, Searching in sub menu: {menu_id}")
                        result, sub = submenu.find_item_pos(cmd, search_sub_menu)
                        if result != -1:
                            if self._logger.is_debug:
                                self._logger.debug(f"Found {str_cmd} in sub menu in position: {result}")
                            return result, sub
                        else:
                            if self._logger.is_debug:
                                self._logger.debug(f"Command {str_cmd} not found in sub menu: {menu_id}")
                menu_type = pop_mnu.get_item_type(i)
                if menu_type == MenuItemType.SEPARATOR:
                    continue
                command = pop_mnu.get_command(menu_id)
                if cmd == command.casefold():
                    if self._logger.is_debug:
                        self._logger.debug(f"Found {str_cmd} in menu: {menu_id}")
                    result = i
                    submenu = pop_mnu
                    return result, submenu
            if result == -1:
                if self._logger.is_debug:
                    self._logger.debug(f"Command {str_cmd} not found.")
                return -1, None
            if self._logger.is_debug:
                self._logger.debug(f"Command {str_cmd} found.")
            return result, submenu

        if cmd.startswith(".custom:"):
            cmd = cmd[8:]
        if not cmd:
            return -1, None
        pos, found_mnu = search(self, cmd)
        if found_mnu is not None:
            found_mnu.cache.capacity = 30
            if self._logger.is_debug:
                self._logger.debug(f"Setting cache capacity back to {found_mnu.cache.capacity} for found menu")
        srch_result = (pos, found_mnu)
        self._cache[key] = srch_result
        return srch_result

    def find_item_menu_id(self, cmd: str, search_sub_menu: bool = False) -> Tuple[int, PopupMenu | None]:
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
        result, submenu = self.find_item_pos(cmd, search_sub_menu)
        if result == -1:
            if self._logger.is_debug:
                self._logger.debug(f"Command {cmd} not found.")
            return -1, None
        if submenu is not None:
            if self._logger.is_debug:
                self._logger.debug(f"Found {cmd} in sub menu. Menu ID: {submenu.get_item_id(result)}")
            return (submenu.get_item_id(result), submenu)  # type: ignore
        if self._logger.is_debug:
            self._logger.debug(f"Found {cmd} in this menu. Menu ID: {self.get_item_id(result)}")
        return (self.get_item_id(result), submenu)  # type: ignore

    # endregion Find methods

    # region MenuPartial Overrides
    def get_popup_menu(self, menu_id: int) -> PopupMenu | None:
        """
        Gets the popup menu from the menu item.

        Args:
            menu_id (int): Menu item id.

        Returns:
            PopupMenu: ``PopupMenu`` instance if found, otherwise ``None``.
        """
        menu = self.component.getPopupMenu(menu_id)
        if menu is None:
            return None
        return PopupMenu(menu)

    def clear(self) -> None:
        """
        Removes all items from the menu.
        """
        super().clear()  # type: ignore
        self._cache.clear()

    def insert_item(self, menu_id: int, text: str, item_style: int | MenuItemStyleKind, item_pos: int) -> None:
        super().insert_item(menu_id=menu_id, text=text, item_style=item_style, item_pos=item_pos)  # type: ignore
        self._cache.clear()

    def remove_item(self, item_pos: int, count: int) -> None:
        super().remove_item(item_pos=item_pos, count=count)  # type: ignore
        self._cache.clear()

    def set_popup_menu(self, menu_id: int, popup_menu: XPopupMenu | ComponentProp) -> None:
        """
        Sets the popup menu for a specified menu item.

        Args:
            menu_id (int): Menu item id.
            popup_menu (XPopupMenu | ComponentProp): Popup menu.
        """
        if isinstance(popup_menu, ComponentProp):
            pm = popup_menu.component
        else:
            pm = popup_menu
        self.component.setPopupMenu(menu_id, pm)

    # endregion MenuPartial Overrides

    # region execute command
    def execute_cmd(self, menu_id: int, in_thread: bool = False) -> bool:
        """
        Executes a command.

        Args:
            cmd (str): Command to execute.
        """
        cmd = self.get_command(menu_id)
        if not cmd:
            return False
        supported_prefixes = tuple(self.lo_inst.get_supported_dispatch_prefixes())
        if cmd.startswith(supported_prefixes):
            self.lo_inst.dispatch_cmd(cmd, in_thread=in_thread)
            return True
        try:
            _ = MacroScript.call_url(cmd, in_thread=in_thread)
            return True
        except Exception as e:
            self._logger.error(f"Error executing menu item with command value of '{cmd}': {e}")
        return False

    # endregion execute command

    # region Properties
    @property
    def cache(self) -> LRUCache:
        """
        Gets the cache.

        Returns:
            LRUCache: Cache.
        """
        return self._cache

    # endregion Properties
