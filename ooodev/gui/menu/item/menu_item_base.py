from __future__ import annotations
from typing import Any, Dict, Tuple
import uno
from com.sun.star.beans import PropertyValue

from ooodev.loader.inst.service import Service
from ooodev.adapter.container.index_access_comp import IndexAccessComp
from ooodev.utils import props as mProps


class MenuItemBase:
    """Base class for individual menu item"""

    def __init__(
        self, *, data: Tuple[Tuple[PropertyValue, ...], ...], owner: IndexAccessComp, app: str | Service = ""
    ):
        """
        Constructor

        Args:
            component (XIndexAccess): UNO Object containing menu item properties.
        """
        self._owner = owner
        self._data = data
        self._menu_data = mProps.Props.data_to_dict(self._data)  # type: ignore
        self._menu_type = int(self._menu_data.get("Type", 0))
        self._app = str(app)

    @property
    def menu_type(self) -> int:
        """Get menu type"""
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

    # @menu_type.setter
    # def menu_type(self, value: MenuTypeKind):
    #     """Set menu type"""
    #     self._menu_type = value
    #     self._menu_data["Type"] = int(value)
