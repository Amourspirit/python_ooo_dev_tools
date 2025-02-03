from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING, Tuple
import contextlib
import threading
from logging import DEBUG

try:
    # python 3.12+
    from typing import override  # noqa # type: ignore
except ImportError:
    from typing_extensions import override  # noqa # type: ignore

from com.sun.star.awt import XPopupMenu
from com.sun.star.frame import XDispatchHelper
from ooo.dyn.awt.menu_item_type import MenuItemType

from ooodev.adapter.component_prop import ComponentProp
from ooodev.io.log.named_logger import NamedLogger
from ooodev.io.log import logging as logger
from ooodev.adapter.awt.popup_menu_comp import PopupMenuComp
from ooodev.utils import gen_util as mGenUtil
from ooodev.utils.cache.singleton.lru_cache import LRUCache
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.macro.script.macro_script import MacroScript
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.beans import PropertyValue
    from com.sun.star.frame import XDispatchProvider
    from com.sun.star.frame import DispatchHelper  # service
    from ooodev.utils.kind.menu_item_style_kind import MenuItemStyleKind
    from ooodev.loader.inst.lo_inst import LoInst


class PopupMenu(LoInstPropsPartial, PopupMenuComp):
    """
    Popup Menu Class.

    See Also:

        - :ref:`help_working_with_menu_bar`
        - :ref:`help_common_gui_menus_menu_bar`
    """

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

        self._cache = LRUCache(capacity=30, name="PopupMenu", key="ooodev.gui.menu.popup_menu.PopupMenu")

    # region Protected Methods

    @override
    def _get_index(self, idx: int, allow_greater: bool = False) -> int:
        """
        Gets the index.

        Args:
            idx (int): Index of sheet. Can be a negative value to index from the end of the list.
            allow_greater (bool, optional): If True and index is greater then the number of
                sheets then the index becomes the next index if sheet were appended. Defaults to False.

        Returns:
            int: Index value.
        """
        count = len(self)
        return mGenUtil.Util.get_index(idx, count, allow_greater)

    def _get_existing_index(self, itm: int | str) -> int:
        """
        Gets the position from the index or command.

        Args:
            item_pos (int | str): Index or command.

        Returns:
            int: Position.
        """
        result = -1
        if isinstance(itm, str):
            if itm == "":
                return result
            result, _ = self.find_item_pos(itm)
        else:
            with contextlib.suppress(ValueError):
                result = self._get_index(itm)
        return result

    # endregion Protected Methods

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
                self._logger.debug("Searching for command: %s", str_cmd)
            cmd = str_cmd.casefold()
            for i, menu_id in enumerate(pop_mnu):
                if search_sub_menu:
                    if self._logger.is_debug:
                        self._logger.debug("Searching %s in sub menu: %i", str_cmd, menu_id)
                    submenu = pop_mnu.get_popup_menu(menu_id)
                    if submenu is not None and len(submenu) > 0:
                        # turn of caching for sub menu when searching. No real value.
                        submenu.cache.capacity = 0
                        if self._logger.is_debug:
                            self._logger.debug("Found a sub Menu %s, Searching in sub menu: %i", str_cmd, menu_id)
                        result, sub = submenu.find_item_pos(cmd, search_sub_menu)
                        if result != -1:
                            if self._logger.is_debug:
                                self._logger.debug("Found %s in sub menu in position: %i", str_cmd, result)
                            return result, sub
                        else:
                            if self._logger.is_debug:
                                self._logger.debug("Command $s not found in sub menu: %i", str_cmd, menu_id)
                menu_type = pop_mnu.get_item_type(i)
                if menu_type == MenuItemType.SEPARATOR:
                    continue
                command = pop_mnu.get_command(menu_id)
                if cmd == command.casefold():
                    if self._logger.is_debug:
                        self._logger.debug("Found %s in menu: %i", str_cmd, menu_id)
                    result = i
                    submenu = pop_mnu
                    return result, submenu
            if result == -1:
                if self._logger.is_debug:
                    self._logger.debug("Command %s not found.", str_cmd)
                return -1, None
            if self._logger.is_debug:
                self._logger.debug("Command %s found.", str_cmd)
            return result, submenu

        if cmd.startswith(".custom:"):
            cmd = cmd[8:]
        if not cmd:
            return -1, None
        pos, found_mnu = search(self, cmd)
        if found_mnu is not None:
            found_mnu.cache.capacity = 30
            if self._logger.is_debug:
                self._logger.debug("Setting cache capacity back to %i for found menu", found_mnu.cache.capacity)
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
                self._logger.debug("Command %s not found.", cmd)
            return -1, None
        if submenu is not None:
            if self._logger.is_debug:
                self._logger.debug("Found %s in sub menu. Menu ID: %i", cmd, submenu.get_item_id(result))
            return (submenu.get_item_id(result), submenu)  # type: ignore
        if self._logger.is_debug:
            self._logger.debug("Found %s in this menu. Menu ID: %i", cmd, self.get_item_id(result))
        return (self.get_item_id(result), submenu)  # type: ignore

    # endregion Find methods

    # region MenuPartial Overrides
    @override
    def get_popup_menu(self, menu_id: int) -> PopupMenu | None:
        """
        Gets the popup menu from the menu item.

        Args:
            menu_id (int): Menu item id.

        Returns:
            PopupMenu: ``PopupMenu`` instance if found, otherwise ``None``.
        """
        menu = self.component.getPopupMenu(menu_id)
        return None if menu is None else PopupMenu(menu)

    @override
    def clear(self) -> None:
        """
        Removes all items from the menu.
        """
        super().clear()  # type: ignore
        self._cache.clear()

    @override
    def insert_item(self, menu_id: int, text: str, item_style: int | MenuItemStyleKind = 0, item_pos: int = -1) -> int:  # type: ignore
        """
        Insert an item into the menu.

        Args:
            menu_id (int): Menu item id. If set to ``-1``, it will be set to the maximum menu id + 1.
            text (str): Text of the menu item.
            item_style (int | MenuItemStyleKind): Style Kind. Defaults to ``0`` (none).
            item_pos (int): Index position of the new item. Can be a negative value to index from the end of the list. Defaults to ``-1``.

        Returns:
            int: Menu item id.
        """
        if menu_id == -1:
            menu_id = self.get_max_menu_id() + 1
        idx = self._get_index(item_pos, allow_greater=True)
        super().insert_item(menu_id=menu_id, text=text, item_style=item_style, item_pos=idx)  # type: ignore
        self._cache.clear()
        return menu_id

    @override
    def insert_item_after(
        self, menu_id: int, text: str, item_style: int | MenuItemStyleKind = 0, after: str | int = -1
    ) -> int:
        """
        Insert an item into the menu.

        Args:
            menu_id (int): Menu item id. If set to ``-1``, it will be set to the maximum menu id + 1.
            text (str): Text of the menu item.
            item_style (int | MenuItemStyleKind): Style Kind. Defaults to ``0`` (none).
            after (str | int): Index of existing item or command to insert after. Can be a negative value to index from the end of the list. Defaults to ``-1``.

        Raises:
            ValueError: ValueError if ``after`` Item not found.

        Returns:
            int: Menu item id.
        """
        if menu_id == -1:
            menu_id = self.get_max_menu_id() + 1
        item_pos = self._get_existing_index(after)
        if item_pos == -1:
            raise ValueError(f"Item not found: {after}")
        super().insert_item(menu_id=menu_id, text=text, item_style=item_style, item_pos=item_pos + 1)  # type: ignore
        self._cache.clear()
        return menu_id

    @override
    def remove_item(self, item_pos: int, count: int) -> None:
        super().remove_item(item_pos=item_pos, count=count)  # type: ignore
        self._cache.clear()

    @override
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
    def is_dispatch_cmd(self, cmd: str | int) -> bool:
        """
        Check if a command is a dispatch command.

        Args:
            cmd (str, int): Command or the Menu id to get the command from that is to be checked.

        Returns:
            bool: ``True`` if it is a dispatch command; Otherwise, ``False``.
        """
        if isinstance(cmd, int):
            cmd = self.get_command(cmd)
        if not cmd:
            return False
        supported_prefixes = self.lo_inst.get_supported_dispatch_prefixes()
        return cmd.startswith(supported_prefixes)

    def execute_cmd(self, cmd: str | int, in_thread: bool = False) -> bool:
        """
        Executes a command.

        Args:
            cmd (str, int): Command or the Menu id to get the command from that is to be executed.

        Returns:
            bool: ``True`` if the command was executed; Otherwise, ``False``.
                if ``in_thread`` is True then it will always return ``True``.

        Note:
            If the command is not a dispatch command that is listed in :meth:`~ooodev.loader.inst.lo_inst.get_supported_dispatch_prefixes`, it will be executed as a Macro URL.
        """
        if isinstance(cmd, int):
            cmd = self.get_command(cmd)
        if not cmd:
            return False
        supported_prefixes = self.lo_inst.get_supported_dispatch_prefixes()
        if cmd.startswith(supported_prefixes):
            self.lo_inst.dispatch_cmd(cmd, in_thread=in_thread)
            return True
        try:
            _ = MacroScript.call_url(cmd, in_thread=in_thread)
            return True
        except Exception as e:
            self._logger.error("Error executing menu item with command value of '%s': %s", cmd, e)
        return False

    def dispatch_cmd(
        self,
        cmd: str | int,
        in_thread: bool = False,
        *props: PropertyValue,
    ) -> bool:
        """
        Executes a command using a new Dispatch Helper.

        Args:
            cmd (str, int): Command or the Menu id to get the command from that is to be executed.
            in_thread (bool, optional): If ``True``, the command will be executed in a new thread. Defaults to ``False``.
            *props (PropertyValue): Additional properties.

        Returns:
            bool: ``True`` if the command was executed; Otherwise, ``False``.
                if ``in_thread`` is True then it will always return ``True``.

        Note:
            Unlike :meth:`execute_cmd`, this method uses a new Dispatch Helper to execute the command.
            There is no check for supported dispatches. Macro URL's are NOT supported.
        """
        if isinstance(cmd, int):
            cmd = self.get_command(cmd)
        if not cmd:
            return False

        self._logger.debug("Dispatching command: %s", cmd)

        helper = cast(
            "DispatchHelper", mLo.Lo.create_instance_mcf(XDispatchHelper, "com.sun.star.frame.DispatchHelper")
        )

        if in_thread:

            def worker(callback, helper: DispatchHelper, provider: XDispatchProvider, cmd: str, props: tuple) -> None:
                result = helper.executeDispatch(provider, cmd, "", 0, props)
                callback(result, cmd)

            def callback(result: Any, cmd: str) -> None:
                # runs after the threaded dispatch is finished
                self._logger.debug("Finished Dispatched in thread: %s", cmd)

            self._logger.debug("Dispatching in thread: %s", cmd)

            t = threading.Thread(
                target=worker, args=(callback, helper, self.lo_inst.desktop.get_active_frame(), cmd, props)
            )
            t.start()
            return True
        else:
            try:
                frame = cast("XDispatchProvider", self.lo_inst.desktop.get_active_frame())
                helper.executeDispatch(frame, cmd, "", 0, props)
            except Exception as e:
                self._logger.error("Error executing menu item with command value of '%s': %s", cmd, e)
                return False
        return True

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
