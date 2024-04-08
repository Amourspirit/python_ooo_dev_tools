from __future__ import annotations
from typing import Any, cast, List, Dict, TYPE_CHECKING
import uno
from ooodev.utils import props as mProps
from ooodev.gui.menu.menu_debug import MenuDebug
from ooodev.gui.menu.menu_base import MenuBase
from ooodev.loader.inst.service import Service

if TYPE_CHECKING:
    from ooodev.adapter.container.index_access_comp import IndexAccessComp
    from ooodev.adapter.ui.ui_configuration_manager_comp import UIConfigurationManagerComp


class Menu:
    """Class for individual menu"""

    def __init__(self, config: UIConfigurationManagerComp, menus: IndexAccessComp[Any], app: str | Service, menu: Any):
        """
        Constructor

        Args:
            config (UIConfigurationManagerComp): Configuration Manager.
            menus (List[Dict[str, str]]): Menus.
            app (str | Service): Name LibreOffice module.
            menu (Any): Particular menu, UNO Object.
        """
        self._config = config
        self._menus = menus
        self._app = str(app)
        self._parent = menu

    def __contains__(self, name):
        """If exists name in menu"""
        exists = False
        for m in self._parent:
            menu = mProps.Props.data_to_dict(m)
            cmd = menu.get("CommandURL", "")
            if name == cmd:
                exists = True
                break
        return exists

    def __getitem__(self, index):
        """Index access"""
        if isinstance(index, int):
            menu = mProps.Props.data_to_dict(self._parent[index])
        else:
            for m in self._parent:
                menu = mProps.Props.data_to_dict(m)
                cmd = menu.get("CommandURL", "")
                if cmd == index:
                    break

        obj = Menu(self._config, self._menus, self._app, menu["ItemDescriptorContainer"])
        return obj

    def debug(self) -> None:
        """Debug menu"""
        MenuDebug()(self._parent)

    def insert(self, menu: Dict[str, Any], after: int | str = "", save: bool = True) -> None:
        """
        Insert new menu.

        Args:
            menu (Dict[str, Any]): New menu data.
            after (int | str, optional): Insert in after menu. Defaults to "".
            save (bool, optional): For persistent save. Defaults to True.
        """
        mb = MenuBase(config=self._config, menus=self._menus, app=self._app)
        mb.insert(self._parent, menu, after)
        if save:
            self._config.component.store()  # type: ignore

    def remove(self, menu: str) -> None:
        """Remove menu

        :param menu: Menu name
        :type menu: str
        """
        mb = MenuBase(config=self._config, menus=self._menus, app=self._app)
        mb.remove(self._parent, menu)
