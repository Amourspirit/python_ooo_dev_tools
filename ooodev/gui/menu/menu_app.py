from __future__ import annotations
from typing import Any, cast, Dict, TYPE_CHECKING, Tuple
import uno
from com.sun.star.beans import PropertyValue
from ooodev.utils import props as mProps
from ooodev.gui.menu.menu_debug import MenuDebug
from ooodev.gui.menu.menu_base import MenuBase
from ooodev.gui.menu.menu import Menu
from ooodev.loader.inst.service import Service
from ooodev.adapter.ui.the_module_ui_configuration_manager_supplier_comp import (
    TheModuleUIConfigurationManagerSupplierComp,
)
from ooodev.adapter.container.index_access_comp import IndexAccessComp


if TYPE_CHECKING:
    from ooodev.adapter.ui.ui_configuration_manager_comp import UIConfigurationManagerComp


class MenuApp:
    """Class for manager menu by LibreOffice module"""

    NODE = "private:resource/menubar/menubar"
    MENUS = {
        "file": ".uno:PickList",
        "picklist": ".uno:PickList",
        "tools": ".uno:ToolsMenu",
        "help": ".uno:HelpMenu",
        "window": ".uno:WindowList",
        "edit": ".uno:EditMenu",
        "view": ".uno:ViewMenu",
        "insert": ".uno:InsertMenu",
        "format": ".uno:FormatMenu",
        "styles": ".uno:FormatStylesMenu",
        "formatstyles": ".uno:FormatStylesMenu",
        "sheet": ".uno:SheetMenu",
        "data": ".uno:DataMenu",
        "table": ".uno:TableMenu",
        "formatform": ".uno:FormatFormMenu",
        "page": ".uno:PageMenu",
        "shape": ".uno:ShapeMenu",
        "slide": ".uno:SlideMenu",
        "slideshow": ".uno:SlideShowMenu",
    }

    def __init__(self, app: str | Service):
        """
        :param app: LibreOffice Module: calc, writer, draw, impress, math, main
        :type app: str
        """
        self._app = str(app)
        self._config = self._get_config()
        self._menus = self._config.get_settings(self.NODE, True)

    def _get_config(self) -> UIConfigurationManagerComp:
        supp = TheModuleUIConfigurationManagerSupplierComp.from_lo()
        return supp.get_ui_configuration_manager(self._app)

    def debug(self):
        """Debug menu"""
        MenuDebug()(self._menus)
        return

    def __contains__(self, name):
        """If exists name in menu"""
        exists = False
        for m in self._menus:
            menu = mProps.Props.data_to_dict(m)
            cmd = menu.get("CommandURL", "")
            if name == cmd:
                exists = True
                break
        return exists

    def __getitem__(self, index):
        """Index access"""
        if isinstance(index, int):
            menu = mProps.Props.data_to_dict(self._menus[index])
        else:
            for m in self._menus:
                menu = mProps.Props.data_to_dict(m)
                cmd = menu.get("CommandURL", "")
                if cmd == index or cmd == self.MENUS[index.lower()]:
                    break
        ia = menu["ItemDescriptorContainer"]
        if ia is None:
            ia_menu = None
        else:
            ia_menu = cast(IndexAccessComp[Tuple[PropertyValue, ...]], IndexAccessComp(ia))
        obj = Menu(self._config, self._menus, self._app, ia_menu)
        return obj

    def insert(self, menu: Dict[str, Any], after: int | str = "", save: bool = True):
        """
        Insert new menu.

        Args:
            menu (dict): New menu data
            after (int, str, optional): Insert in after menu. Defaults to "".
            save (bool, optional): For persistent save. Defaults to True.
        """
        mb = MenuBase(config=self._config, menus=self._menus, app=self._app)

        mb.insert(self._menus, menu, after)
        if save:
            self._config.component.store()  # type: ignore
        return

    def remove(self, menu: str) -> None:
        """
        Remove menu.

        Args:
            menu (str): Menu name
        """
        mb = MenuBase(config=self._config, menus=self._menus, app=self._app)
        mb.remove(self._menus, menu)
        return
