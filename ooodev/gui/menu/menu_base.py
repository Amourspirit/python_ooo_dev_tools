from __future__ import annotations
from typing import Any, Union, List, Dict, TYPE_CHECKING, Tuple
import uno
from com.sun.star.beans import PropertyValue

from ooodev.adapter.container.index_access_comp import IndexAccessComp
from ooodev.gui.menu.shortcuts import Shortcuts
from ooodev.io.log.logging import debug, error
from ooodev.loader import lo as mLo
from ooodev.loader.inst.service import Service
from ooodev.macro.script.macro_script import MacroScript
from ooodev.utils import props as mProps
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial

if TYPE_CHECKING:
    from com.sun.star.container import XIndexContainer
    from ooodev.adapter.ui.ui_configuration_manager_comp import UIConfigurationManagerComp
    from ooodev.loader.inst.lo_inst import LoInst


class MenuBase(LoInstPropsPartial):
    """Class Base for menus"""

    NODE = "private:resource/menubar/menubar"

    def __init__(
        self,
        *,
        config: UIConfigurationManagerComp,
        menus: IndexAccessComp[Tuple[PropertyValue, ...]],
        app: str | Service = "",
        lo_inst: LoInst | None = None,
    ):
        """
        Constructor

        Args:
            app (str | Service, optional): App Name such as ``Service.CALC``. Defaults to "".
                If no app is provided, the global shortcuts will be used.

        Hint:
            - ``Service`` is an enum and can be imported from ``ooodev.loader.inst.service``
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        LoInstPropsPartial.__init__(self, lo_inst)
        self._app = str(app)
        self._config = config
        self._menus = menus

    def _get_index(self, parent: Any, name: Union[int, str] = "") -> int:
        """
        Get index menu from name.

        Args:
            parent (Any): Menu parent, UNO object.
            name (Union[int, str], optional): Menu name. Defaults to "".

        Returns:
            int: Index of menu
        """
        index = -1
        if isinstance(name, str) and name:
            for i, m in enumerate(parent):
                menu = mProps.Props.data_to_dict(m)
                if menu.get("CommandURL", "") == name:
                    index = i
                    break
        elif isinstance(name, str):
            index = len(parent) - 1
        elif isinstance(name, int):
            index = name
        return index

    def _get_command_url(self, menu: Dict[str, str]):
        """
        Get url from command and set shortcut.

        Args:
            menu (dict): Menu data

        Returns:
            str: URL command
        """
        shortcut = menu.pop("ShortCut", "")
        command = menu["CommandURL"]
        url = Shortcuts.get_url_script(command)
        if shortcut:
            Shortcuts(self._app).set(shortcut, command)
        return url

    def _save(self, parent: Any, menu: Dict[str, Any], index: int):
        """
        Insert menu.

        Args:
            parent (Any): Menu parent
            menu (dict): Menu data
            index (int): Position to insert

        Returns:
            None:
        """
        properties = mProps.Props.make_props_any(**menu)
        if hasattr(parent, "component"):
            obj = parent.component
        else:
            obj = parent
        uno.invoke(obj, "insertByIndex", (index, properties))  # type: ignore
        self._config.replace_settings(self.NODE, self._menus.component)
        return

    def _insert_submenu(self, parent: XIndexContainer, menus: List[Dict[str, Any]]) -> None:
        """
        Insert submenus recursively

        Args:
            parent (Any): Menu parent
            menus (list): List of menus
        """
        for i, menu in enumerate(menus):
            submenu = menu.pop("Submenu", False)
            if submenu:
                idc = self._config.component.createSettings()
                menu["ItemDescriptorContainer"] = idc
            menu["Type"] = 0
            if menu["Label"][0] == "-":
                menu["Type"] = 1
            else:
                menu["CommandURL"] = self._get_command_url(menu)
            self._save(parent, menu, i)
            if submenu:
                self._insert_submenu(idc, submenu)

    def _get_first_command(self, command: str | Dict[str, str]):
        url = command
        if isinstance(command, dict):
            url = MacroScript.get_url_script(**command)
        return url

    def insert(self, parent: Any, menu: Dict[str, Any], after: int | str = "") -> None:
        """
        Insert new menu.

        Args:
            parent (Any): Menu parent
            menu (Dict[str, Any]): New menu data
            after (int | str, optional): After menu insert. Defaults to "".
        """
        parent_index = self._get_index(parent, after)
        if parent_index == -1:
            msg = f"Parent not found: {after}"
            error(msg)
            raise ValueError(msg)
        index = parent_index + 1
        submenu = menu.pop("Submenu", False)
        menu["Type"] = 0
        idc = self._config.component.createSettings()
        menu["ItemDescriptorContainer"] = idc
        menu["CommandURL"] = self._get_first_command(menu["CommandURL"])
        self._save(parent, menu, index)
        if submenu:
            self._insert_submenu(idc, submenu)

    def remove(self, parent: Any, name: str | Dict[str, str]) -> None:
        """
        Remove name in parent.

        Args:
            parent (Any): Menu parent. UNO object.
            name (str,  Dict[str, str]): Menu CommandURL or data

        Returns:
            None:
        """
        if isinstance(name, dict):
            name = MacroScript.get_url_script(**name)
        index = self._get_index(parent, name)
        if index == -1:
            debug(f"Not found: {name}")
            return
        uno.invoke(parent, "removeByIndex", (index,))  # type: ignore
        self._config.replace_settings(self.NODE, self._menus.component)
        self._config.component.store()  # type: ignore

    @property
    def app(self) -> str:
        return self._app

    @property
    def config(self) -> UIConfigurationManagerComp:
        return self._config

    @property
    def menus(self) -> IndexAccessComp[Tuple[PropertyValue, ...]]:
        return self._menus
