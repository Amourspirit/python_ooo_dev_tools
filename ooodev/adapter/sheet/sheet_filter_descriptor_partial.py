from __future__ import annotations
from typing import Any, TYPE_CHECKING, Tuple
import uno
from com.sun.star.sheet import XSheetFilterDescriptor

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.sheet import TableFilterField
    from com.sun.star.sheet import XSheetFilterDescriptor
    from ooodev.utils.type_var import UnoInterface


class SheetFilterDescriptorPartial:
    """
    Partial Class for XSheetFilterDescriptor.
    """

    def __init__(
        self, component: XSheetFilterDescriptor, interface: UnoInterface | None = XSheetFilterDescriptor
    ) -> None:
        """
        Constructor

        Args:
            component (XSheetFilterDescriptor): UNO Component that implements ``com.sun.star.sheet.XSheetFilterDescriptor``.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XSheetFilterDescriptor``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XSheetFilterDescriptor

    def get_filter_fields(self) -> Tuple[TableFilterField, ...]:
        """
        returns the collection of filter fields.
        """
        return self.__component.getFilterFields()

    def set_filter_fields(self, *filter_fields: TableFilterField) -> None:
        """
        Sets a new collection of filter fields.
        """
        self.__component.setFilterFields(filter_fields)

    # endregion XDatabaseRange
