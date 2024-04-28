from __future__ import annotations
from typing import TYPE_CHECKING

from ooodev.adapter.container.index_access_comp import IndexAccessComp
from ooodev.gui.menu.menu import Menu
from ooodev.gui.menu.ma.ma_popup import MAPopup
from ooodev.loader.inst.service import Service
from ooodev.utils import props as mProps
from ooodev.utils.kind.menu_lookup_kind import MenuLookupKind

if TYPE_CHECKING:
    from ooodev.loader.inst.lo_inst import LoInst


class MenuApp(MAPopup):
    """
    Class for manager menu by LibreOffice module.

    See Also:
        - :ref:`help_creating_menu_using_menu_app`
        - :ref:`help_working_with_menu_app`
    """

    NODE = "private:resource/menubar/menubar"
    MENUS = MenuLookupKind.get_dict()

    def __init__(self, app: str | Service, lo_inst: LoInst | None = None):
        """
        Constructor

        Args:
            app (str | Service): LibreOffice Module: calc, writer, draw, impress, math, main
            lo_inst (LoInst | None, optional): LibreOffice instance. Defaults to ``None``.
        """
        super().__init__(app=app, node=self.NODE, lo_inst=lo_inst)

    def __getitem__(self, index: int | str | MenuLookupKind) -> Menu:
        """
        Index access.

        Args:
            index (int, str, MenuLookupKind): Index or Menu name or MenuLookupKind or CommandURL.
                If index is a str then it can be a known menu name or a CommandURL.

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
        menus = self._get_menus()
        cache = self._get_cache()
        cache_key = f"get_item_{index}"
        if cache_key in self._cache:
            return cache[cache_key]
        menu = None
        if isinstance(index, int):
            if index < 0 or index >= len(menus):
                raise IndexError(f"Index out of range: {index}")
            menu = mProps.Props.data_to_dict(menus[index])
        else:
            if isinstance(index, str):
                key = self.MENUS[index.lower()] if index.lower() in self.MENUS else index
            else:
                # MenuLookupKind
                # By calling str() on the enum, we get the value of the enum.
                # This can also allow other custom enums to be used as lookup values in the future.
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
            node=self.NODE,
            config=self._get_config(),
            menus=menus,
            app=self.app,
            menu=ia_menu,  # type: ignore
            lo_inst=self.lo_inst,
        )
        cache[cache_key] = obj
        return obj
