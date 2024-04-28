from __future__ import annotations
from typing import Any, cast, Dict, TYPE_CHECKING
from ooodev.adapter.container.index_access_comp import IndexAccessComp
from ooodev.adapter.container.index_access_implement import IndexAccessImplement
from ooodev.adapter.ui.the_module_ui_configuration_manager_supplier_comp import (
    TheModuleUIConfigurationManagerSupplierComp,
)
from ooodev.gui.menu.menu import Menu
from ooodev.gui.menu.common.menu_base import MenuBase
from ooodev.gui.menu.common.menu_debug import MenuDebug
from ooodev.loader import lo as mLo
from ooodev.loader.inst.service import Service
from ooodev.utils import props as mProps
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.cache.lru_cache import LRUCache

if TYPE_CHECKING:
    from com.sun.star.container import XIndexAccess
    from ooodev.adapter.ui.ui_configuration_manager_comp import UIConfigurationManagerComp
    from ooodev.loader.inst.lo_inst import LoInst


# https://opengrok.libreoffice.org/xref/core/officecfg/registry/data/org/openoffice/Office/UI/


class MAPopup(LoInstPropsPartial):
    """
    Class for manager menu by LibreOffice module.

    See Also:
        - :ref:`help_creating_menu_using_menu_app`
        - :ref:`help_working_with_menu_app`
    """

    def __init__(self, app: str | Service, node: str, lo_inst: LoInst | None = None):
        """
        Constructor

        Args:
            app (str | Service): LibreOffice Module: calc, writer, draw, impress, math, main
            node (str): Menu Node such as ``private:resource/menubar/menubar``
            lo_inst (LoInst | None, optional): LibreOffice instance. Defaults to ``None``.
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        self._node = node
        LoInstPropsPartial.__init__(self, lo_inst)
        self._app = str(app)
        self._cache = LRUCache(50)

    def _get_cache(self) -> LRUCache:
        return self._cache

    def _get_config(self) -> UIConfigurationManagerComp:
        cache = self._get_cache()
        key = "_get_config"
        if key in cache:
            return cache[key]
        supp = TheModuleUIConfigurationManagerSupplierComp.from_lo(lo_inst=self.lo_inst)
        cm = supp.get_ui_configuration_manager(self.app)
        cache[key] = cm
        return cast("UIConfigurationManagerComp", cm)

    def _get_settings(self) -> IndexAccessComp[Any]:
        return self._get_config().get_settings(self._node, True)

    def _get_menus(self) -> IndexAccessComp[Any]:
        cache = self._get_cache()
        key = "_get_menus"
        if key in cache:
            return cache[key]
        menus = self._get_settings()
        cache[key] = menus
        return menus

    def _get_empty_index_access(self) -> XIndexAccess:
        return IndexAccessImplement(elements=(), element_type="[]com.sun.star.beans.PropertyValue")

    def debug(self):
        """Debug menu"""
        MenuDebug()(self._get_menus())
        return

    def __contains__(self, name: str):
        """
        If exists name in menu.

        Args:
            name (str): Menu CommandURL.
        """
        cache = self._get_cache()
        key = f"contains_{name}"
        if key in cache:
            return cache[key]
        exists = False
        for m in self._get_menus():
            menu = mProps.Props.data_to_dict(m)
            cmd = menu.get("CommandURL", "")
            if name == cmd:
                exists = True
                break
        cache[key] = exists
        return exists

    def __getitem__(self, index: Any) -> Menu:
        """
        Index access.

        Args:
            index (Any): Index or CommandURL or obj the can convert to str.

        Raises:
            IndexError: Index out of range.
            KeyError: Menu not found.

        Returns:
            Menu: Menu instance.

        Hint:
            - ``MenuLookupKind`` is an enum and can be imported from ``ooodev.utils.kind.menu_lookup_kind``

        Note:
            Index can also be any object the returns a command URL when str() is called on it.
        """
        cache = self._get_cache()
        cache_key = f"get_item_{index}"
        if cache_key in cache:
            return cache[cache_key]
        menu = None
        menus = self._get_menus()
        if isinstance(index, int):
            if index < 0 or index >= len(menus):
                raise IndexError(f"Index out of range: {index}")
            menu = mProps.Props.data_to_dict(menus[index])
        else:
            key = str(index)
            for m in menus:
                menu = mProps.Props.data_to_dict(m)
                cmd = menu.get("CommandURL", "")
                if cmd == key:
                    break
        if menu is None:
            raise KeyError(f"Menu not found: {index}")
        ia = menu.get("ItemDescriptorContainer", None)
        if ia is None:
            # create an empty XIndexAccess
            ia = self._get_empty_index_access()

        ia_menu = IndexAccessComp(ia)
        obj = Menu(
            node=self._node,
            config=self._get_config(),
            menus=menus,
            app=self._app,
            menu=ia_menu,  # type: ignore
            lo_inst=self.lo_inst,
        )
        cache[cache_key] = obj
        return obj

    def insert(self, menu: Dict[str, Any], after: int | str = "", save: bool = True):
        """
        Insert new menu.

        Args:
            menu (dict): New menu data
            after (int, str, optional): Insert in after menu (CommandURL). Defaults to "".
            save (bool, optional): For persistent save. Defaults to True.
        """
        menus = self._get_menus()
        config = self._get_config()
        mb = MenuBase(node=self._node, config=self._get_config(), menus=menus, app=self._app, lo_inst=self.lo_inst)

        mb.insert(menus, menu, after)
        if save:
            config.component.store()  # type: ignore
        return

    def remove(self, menu: str) -> None:
        """
        Remove menu.

        Args:
            menu (str): Menu CommandURL
        """
        menus = self._get_menus()
        cache = self._get_cache()
        mb = MenuBase(node=self._node, config=self._get_config(), menus=menus, app=self._app, lo_inst=self.lo_inst)
        mb.remove(menus, menu)
        cache.clear()
        return

    @property
    def app(self) -> str:
        """Gets the current app."""
        return self._app
