from __future__ import annotations
from typing import Any, Dict, Tuple, TYPE_CHECKING
import uno
from com.sun.star.beans import PropertyValue

from ooodev.adapter.container.index_access_comp import IndexAccessComp
from ooodev.loader.inst.service import Service
from ooodev.loader import lo as mLo
from ooodev.utils import props as mProps
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial

if TYPE_CHECKING:
    from ooodev.gui.menu.menu import Menu
    from ooodev.loader.inst.lo_inst import LoInst


class MenuItemBase(LoInstPropsPartial):
    """Base class for individual menu item"""

    def __init__(
        self,
        *,
        menu: Menu,
        data: Tuple[Tuple[PropertyValue, ...], ...],
        owner: IndexAccessComp,
        app: str | Service = "",
        lo_inst: LoInst | None = None,
    ):
        """
        Constructor

        Args:
            component (XIndexAccess): UNO Object containing menu item properties.
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        LoInstPropsPartial.__init__(self, lo_inst)
        self._menu = menu
        self._owner = owner
        self._data = data
        self._menu_data = mProps.Props.data_to_dict(self._data)  # type: ignore
        self._menu_type = int(self._menu_data.get("Type", 0))
        self._app = str(app)

    def __bool__(self) -> bool:
        return self._data is not None

    @property
    def menu_type(self) -> int:
        """
        Get menu type.

        Returns:
            int: Menu type. ``0`` for ``MenuItem``, ``1`` for ``MenuItemSep``.
        """
        return self._menu_type

    @property
    def data(self) -> Tuple[Tuple[PropertyValue, ...], ...]:
        """Get menu data"""
        return self._data

    @property
    def data_dict(self) -> Dict[str, Any]:
        """Get menu data as dictionary"""
        return self._menu_data

    @property
    def app(self) -> str:
        """Get app"""
        return self._app

    @property
    def menu(self) -> Menu:
        """Get menu that owns this item."""
        return self._menu

    # @menu_type.setter
    # def menu_type(self, value: MenuTypeKind):
    #     """Set menu type"""
    #     self._menu_type = value
    #     self._menu_data["Type"] = int(value)
