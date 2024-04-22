from __future__ import annotations
from typing import Any, Union, List, Dict, TYPE_CHECKING, Tuple, Iterable
import uno
from com.sun.star.beans import PropertyValue

from ooodev.adapter.component_prop import ComponentProp
from ooodev.adapter.container.index_access_comp import IndexAccessComp
from ooodev.gui.menu.shortcuts import Shortcuts
from ooodev.io.log.named_logger import NamedLogger
from ooodev.loader import lo as mLo
from ooodev.loader.inst.service import Service
from ooodev.utils import props as mProps
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial

if TYPE_CHECKING:
    from com.sun.star.container import XIndexContainer
    from ooodev.adapter.ui.ui_configuration_manager_comp import UIConfigurationManagerComp
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.gui.menu.common.command_dict import CommandDict


class MenuBase(LoInstPropsPartial):
    """Class Base for menus"""

    VALID_KEYS = {"Label", "CommandURL", "ShortCut", "Style", "Type", "ItemDescriptorContainer"}

    def __init__(
        self,
        *,
        node: str,
        config: UIConfigurationManagerComp,
        menus: IndexAccessComp[Tuple[PropertyValue, ...]],
        app: str | Service = "",
        lo_inst: LoInst | None = None,
    ):
        """
        Constructor

        Args:
            node (str): Menu Node such as ``private:resource/menubar/menubar``.
            config (UIConfigurationManagerComp): Configuration Manager.
            menus (IndexAccessComp[Tuple[PropertyValue, ...]]): Menus.
            app (str | Service, optional): App Name such as ``Service.CALC``. Defaults to "".
                If no app is provided, the global shortcuts will be used.
            lo_inst (LoInst, optional): LibreOffice Instance. Defaults to None.

        Hint:
            - ``Service`` is an enum and can be imported from ``ooodev.loader.inst.service``
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        self._node = node
        LoInstPropsPartial.__init__(self, lo_inst)
        self._app = str(app)
        self._config = config
        self._menus = menus
        self._logger = NamedLogger(self.__class__.__name__)

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
        save = True
        sc = menu.pop("ShortCut", None)
        if sc is None:
            shortcut = ""
        elif isinstance(sc, str):
            shortcut = sc
        elif isinstance(sc, dict):
            shortcut = sc.get("key", "")
            save = sc.get("save", False)
        else:
            self._logger.error(f"Invalid shortcut: {sc}")
            shortcut = ""

        command = menu["CommandURL"]
        url = Shortcuts.get_url_script(command)
        if shortcut:
            Shortcuts(self._app).set(shortcut, command, save)
        return url

    def validate_keys(self, dictionary: dict, valid_keys: Iterable[str]) -> bool:
        """
        Validate keys in dictionary.

        Args:
            dictionary (dict): Dictionary to validate.
            valid_keys (Iterable[str]): Valid keys.

        Returns:
            bool: True if all keys are valid.
        """
        return all(key in valid_keys for key in dictionary)

    def _process_menu(self, menu: Dict[str, Any]):
        """
        Process menu data.

        Args:
            menu (dict): Menu data

        Returns:
            dict: Processed menu data
        """
        if not self.validate_keys(menu, MenuBase.VALID_KEYS):
            raise ValueError(f"Invalid keys: {menu}")
        mnu = menu.copy()
        if "Style" in mnu:
            mnu["Style"] = int(mnu["Style"])
        return mnu

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
        # changes: are saved in ..\AppData\Roaming\LibreOffice\4\user\config\soffice.cfg\modules\...
        # for calc that would be ..\AppData\Roaming\LibreOffice\4\user\config\soffice.cfg\modules\scalc\menubar
        mnu = self._process_menu(menu)
        properties = mProps.Props.make_props_any(**mnu)
        obj = parent.component if hasattr(parent, "component") else parent
        uno.invoke(obj, "insertByIndex", (index, properties))  # type: ignore
        self._config.replace_settings(self._node, self._menus.component)
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

    def _get_first_command(self, command: Union[str, CommandDict]):
        return Shortcuts.get_url_script(command)
        # url = command
        # if isinstance(command, dict):
        #     url = MacroScript.get_url_script(**command)
        # elif isinstance(command, str):
        #     if command.startswith(".custom:"):
        #         url = command[8:]
        # return url

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
            self._logger.error(msg)
            raise ValueError(msg)
        index = parent_index + 1
        submenu = menu.pop("Submenu", False)
        menu["Type"] = 0
        menu["CommandURL"] = self._get_first_command(menu["CommandURL"])
        if submenu:
            idc = self._config.component.createSettings()
            menu["ItemDescriptorContainer"] = idc
        self._save(parent, menu, index)
        if submenu:
            self._insert_submenu(idc, submenu)

    def remove(self, parent: Any, name: Union[str, CommandDict], save: bool = False) -> None:
        """
        Remove name in parent.

        Args:
            parent (Any): Menu parent. UNO object.
            name (str,  Dict[str, str]): Menu ``CommandURL`` or data.
            save (bool, optional): Save changes. Defaults to ``False``.

        Returns:
            None:
        """
        name = Shortcuts.get_url_script(name)
        # if isinstance(name, dict):
        #     name = MacroScript.get_url_script(**name)
        index = self._get_index(parent, name)
        if index == -1:
            self._logger.debug(f"Not found: {name}")
            return
        if isinstance(parent, ComponentProp):
            parent = parent.component
        uno.invoke(parent, "removeByIndex", (index,))  # type: ignore
        self._config.replace_settings(self._node, self._menus.component)
        if save:
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
