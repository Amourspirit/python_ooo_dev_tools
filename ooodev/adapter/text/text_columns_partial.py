from __future__ import annotations
from typing import Any, TYPE_CHECKING, Tuple
import uno

from com.sun.star.text import XTextColumns
from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface
    from com.sun.star.text import TextColumn  # struct


class TextColumnsPartial:
    """
    Partial class for XTextColumns.
    """

    # A text column is a UNO TextColumn struct.

    # pylint: disable=unused-argument

    def __init__(self, component: XTextColumns, interface: UnoInterface | None = XTextColumns) -> None:
        """
        Constructor

        Args:
            component (XTextColumns): UNO Component that implements ``com.sun.star.text.XTextColumns`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XTextColumns``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XTextColumns
    def get_column_count(self) -> int:
        """
        Gets the number of columns.
        """
        return self.__component.getColumnCount()

    def get_columns(self) -> Tuple[TextColumn, ...]:
        """
        returns the column description of the object.
        """
        return self.__component.getColumns()

    def get_reference_value(self) -> int:
        """
        Gets the sum of values, the width values are relative.

        Returns:
            int: The sum of all values.
        """
        return self.__component.getReferenceValue()

    def set_column_count(self, columns: int) -> None:
        """
        Sets the number of columns.

        The minimum is 1 column.
        """
        self.__component.setColumnCount(columns)

    def set_columns(self, *columns: TextColumn) -> None:
        """
        Sets the descriptors of all columns.

        The number of members in the sequence must be the same as the number of columns of the object.
        """
        self.__component.setColumns(columns)

    # endregion XTextColumns
