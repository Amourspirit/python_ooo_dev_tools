from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno
from com.sun.star.sheet import XDatabaseRanges

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.adapter.container.name_access_partial import NameAccessPartial

if TYPE_CHECKING:
    from com.sun.star.table import CellRangeAddress
    from ooodev.utils.type_var import UnoInterface


class DatabaseRangesPartial(NameAccessPartial):
    """
    Partial Class for XDatabaseRanges.
    """

    def __init__(self, component: XDatabaseRanges, interface: UnoInterface | None = XDatabaseRanges) -> None:
        """
        Constructor

        Args:
            component (XDatabaseRanges): UNO Component that implements ``com.sun.star.sheet.XDatabaseRanges``.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XDatabaseRanges``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        NameAccessPartial.__init__(self, component=component, interface=None)
        self.__component = component

    # region XDatabaseRanges
    def add_new_by_name(self, name: str, rng: CellRangeAddress) -> None:
        """
        Adds a new database range to the collection.
        """
        self.__component.addNewByName(name, rng)

    def remove_by_name(self, name: str) -> None:
        """
        Removes a database range from the collection.
        """
        self.__component.removeByName(name)

    # endregion XDatabaseRanges
