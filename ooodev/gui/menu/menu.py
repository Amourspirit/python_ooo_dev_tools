from __future__ import annotations
from typing import Any, Dict, TYPE_CHECKING, Tuple
import uno
from com.sun.star.beans import PropertyValue

from ooodev.adapter.container.index_access_comp import IndexAccessComp
from ooodev.adapter.container.index_access_implement import IndexAccessImplement
from ooodev.gui.menu.item.menu_items import MenuItems
from ooodev.gui.menu.common.menu_base import MenuBase
from ooodev.gui.menu.common.menu_debug import MenuDebug
from ooodev.gui.menu.shortcuts import Shortcuts
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
        *,
        node: str,
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
        self._node = node
        self._config = config
        self._menus = menus
        self._app = str(app)
        self._current_menu = menu
        self._items = MenuItems(component=menu.component, menu=self, app=self._app, lo_inst=self.lo_inst)
        self._current_index = -1

    def __bool__(self) -> bool:
        """
        If menu exists.

        Returns:
            bool: ``True`` if menu exists and menu is not empty.
        """
        return self._current_menu is not None and len(self._current_menu) > 0

    def __contains__(self, name: str | Dict[str, str]) -> bool:
        """If exists name in menu"""
        name = Shortcuts.get_url_script(name)
        exists = False
        for m in self._current_menu:
            menu = mProps.Props.data_to_dict(m)
            cmd = menu.get("CommandURL", "")
            if name == cmd:
                exists = True
                break
        return exists

    def __getitem__(self, index) -> Menu:
        """
        Index access.

        Args:
            index (int | str): Index or CommandURL.

        Raises:
            IndexError: Index out of range.
            KeyError: Menu not found.

        Returns:
            Menu: Menu object.
        """
        menu = None
        if isinstance(index, int):
            count = len(self)
            if index < 0 or index >= count:
                raise IndexError(f"Index out of range: {index}")
            menu = mProps.Props.data_to_dict(self._current_menu[index])
        else:
            name = Shortcuts.get_url_script(index)
            for m in self._current_menu:
                menu = mProps.Props.data_to_dict(m)
                cmd = menu.get("CommandURL", "")
                if cmd == name:
                    break
        if menu is None:
            raise KeyError(f"Menu not found: {index}")
        ia = menu.get("ItemDescriptorContainer", None)
        # if ia is None:
        #     ia_menu = None
        if ia is None:
            # create an empty XIndexAccess
            ia = IndexAccessImplement(elements=(), element_type="[]com.sun.star.beans.PropertyValue")
        ia_menu = IndexAccessComp(ia)
        return Menu(
            node=self._node,
            config=self._config,
            menus=self._menus,
            app=self._app,
            menu=ia_menu,  # type: ignore
            lo_inst=self.lo_inst,
        )

    def __len__(self) -> int:
        """Length of menu"""
        return len(self._current_menu)

    def __iter__(self):
        """Iterator"""
        self._current_index = 0
        return self

    def __next__(self):
        """Next"""
        if self._current_index < len(self):
            menu = self[self._current_index]
            self._current_index += 1
            return menu
        self._current_index = -1
        raise StopIteration

    def __repr__(self) -> str:
        """String representation"""
        # if self._current_menu is not None and len(self._current_menu) > 0:

        #     return ""
        return object.__repr__(self)
        # return f"{self._current_menu}"

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
        mb = MenuBase(node=self._node, config=self._config, menus=self._menus, app=self._app, lo_inst=self.lo_inst)
        mb.insert(self._current_menu, menu, after)
        if save:
            self._config.component.store()  # type: ignore

    def remove(self, menu: str, save: bool = False) -> None:
        """
        Remove menu.

        Args:
            menu (str): Menu name.
            save (bool, optional): Save changes. Defaults to ``False``.
        """
        mb = MenuBase(node=self._node, config=self._config, menus=self._menus, app=self._app, lo_inst=self.lo_inst)
        mb.remove(self._current_menu, menu, save)

    @property
    def items(self) -> MenuItems:
        return self._items
