from __future__ import annotations
from typing import Any, TYPE_CHECKING, Tuple
import uno
from com.sun.star.sheet import XDatabaseRange

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.beans import PropertyValue
    from com.sun.star.table import CellRangeAddress
    from com.sun.star.sheet import XSheetFilterDescriptor
    from com.sun.star.sheet import XSubTotalDescriptor
    from ooodev.utils.type_var import UnoInterface


class DatabaseRangePartial:
    """
    Partial Class for XDatabaseRange.
    """

    def __init__(self, component: XDatabaseRange, interface: UnoInterface | None = XDatabaseRange) -> None:
        """
        Constructor

        Args:
            component (XDatabaseRange): UNO Component that implements ``com.sun.star.sheet.XDatabaseRange``.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XDatabaseRange``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XDatabaseRange

    def get_data_area(self) -> CellRangeAddress:
        """
        Returns the data area of the database range in the spreadsheet document.
        """
        return self.__component.getDataArea()

    def get_filter_descriptor(self) -> XSheetFilterDescriptor:
        """
        Returns the filter descriptor stored with the database range.

        If the filter descriptor is modified, the new filtering is carried out when XDatabaseRange.refresh() is called.
        """
        return self.__component.getFilterDescriptor()

    def get_import_descriptor(self) -> Tuple[PropertyValue, ...]:
        """
        Returns the database import descriptor stored with this database range.
        """
        return self.__component.getImportDescriptor()

    def get_sort_descriptor(self) -> Tuple[PropertyValue, ...]:
        """
        Returns the sort descriptor stored with the database range.
        """
        return self.__component.getSortDescriptor()

    def get_sub_total_descriptor(self) -> XSubTotalDescriptor:
        """
        Returns the subtotal descriptor stored with the database range.

        If the subtotal descriptor is modified, the new subtotals are inserted when XDatabaseRange.refresh() is called.
        """
        return self.__component.getSubTotalDescriptor()

    def refresh(self) -> None:
        """
        Executes the stored import, filter, sorting, and subtotals descriptors again.
        """
        self.__component.refresh()

    def set_data_area(self, data_area: CellRangeAddress) -> None:
        """
        Sets the data area of the database range.
        """
        self.__component.setDataArea(data_area)

    # endregion XDatabaseRange
