from __future__ import annotations
from typing import Any, Dict, TYPE_CHECKING, Tuple
import uno
from com.sun.star.beans import PropertyValue

from ooodev.adapter.container.index_access_comp import IndexAccessComp
from ooodev.gui.menu.item.menu_items import MenuItems
from ooodev.gui.menu.menu_base import MenuBase
from ooodev.gui.menu.menu_debug import MenuDebug
from ooodev.loader import lo as mLo
from ooodev.loader.inst.service import Service
from ooodev.utils import props as mProps
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial

if TYPE_CHECKING:
    from ooodev.adapter.ui.ui_configuration_manager_comp import UIConfigurationManagerComp
    from ooodev.loader.inst.lo_inst import LoInst


class Menu(LoInstPropsPartial):
    """Class for individual menu"""

    def __init__(
        self,
        config: UIConfigurationManagerComp,
        menus: IndexAccessComp[Tuple[PropertyValue, ...]],
        app: str | Service,
        menu: IndexAccessComp[Tuple[PropertyValue, ...]],
        lo_inst: LoInst | None = None,
    ) -> None:
        """
        Constructor

        Args:
            config (UIConfigurationManagerComp): Configuration Manager.
            menus (List[Dict[str, str]]): Menus.
            app (str | Service): Name LibreOffice module.
            menu (Any): Particular menu, UNO Object.
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        LoInstPropsPartial.__init__(self, lo_inst)
        self._config = config
        self._menus = menus
        self._app = str(app)
        self._current_menu = menu
        self._items = MenuItems(component=menu.component, lo_inst=self.lo_inst)

    def __contains__(self, name) -> bool:
        """If exists name in menu"""
        exists = False
        for m in self._current_menu:
            menu = mProps.Props.data_to_dict(m)
            cmd = menu.get("CommandURL", "")
            if name == cmd:
                exists = True
                break
        return exists

    def __getitem__(self, index):
        """Index access"""
        if isinstance(index, int):
            menu = mProps.Props.data_to_dict(self._current_menu[index])
        else:
            for m in self._current_menu:
                menu = mProps.Props.data_to_dict(m)
                cmd = menu.get("CommandURL", "")
                if cmd == index:
                    break
        ia = menu["ItemDescriptorContainer"]
        if ia is None:
            ia_menu = None
        else:
            ia_menu = IndexAccessComp(ia)
        obj = Menu(
            config=self._config,
            menus=self._menus,
            app=self._app,
            menu=ia_menu,  # type: ignore
            lo_inst=self.lo_inst,
        )
        return obj

    def debug(self) -> None:
        """Debug menu"""
        MenuDebug()(self._current_menu)

    def insert(self, menu: Dict[str, Any], after: int | str = "", save: bool = True) -> None:
        """
        Insert new menu.

        Args:
            menu (Dict[str, Any]): New menu data.
            after (int | str, optional): Insert in after menu. Defaults to "".
            save (bool, optional): For persistent save. Defaults to True.
        """
        mb = MenuBase(config=self._config, menus=self._menus, app=self._app, lo_inst=self.lo_inst)
        mb.insert(self._current_menu, menu, after)
        if save:
            self._config.component.store()  # type: ignore

    def remove(self, menu: str) -> None:
        """
        Remove menu.

        Args:
            menu (str): Menu name.
        """
        mb = MenuBase(config=self._config, menus=self._menus, app=self._app, lo_inst=self.lo_inst)
        mb.remove(self._current_menu, menu)

    @property
    def items(self) -> MenuItems:
        return self._items
