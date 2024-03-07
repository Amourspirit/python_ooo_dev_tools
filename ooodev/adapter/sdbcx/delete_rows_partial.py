from __future__ import annotations
from typing import Any, TYPE_CHECKING, Tuple
import uno

from com.sun.star.sdbcx import XDeleteRows

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class DeleteRowsPartial:
    """
    Partial class for XDeleteRows.
    """

    def __init__(self, component: XDeleteRows, interface: UnoInterface | None = XDeleteRows) -> None:
        """
        Constructor

        Args:
            component (XDeleteRows): UNO Component that implements ``com.sun.star.container.XDeleteRows`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XDeleteRows``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XDeleteRows
    def delete_rows(self, rows: Tuple[Any, ...]) -> Tuple[int, ...]:
        """
        Deletes one or more rows identified by their bookmarks.

        Args:
            rows (Tuple[Any, ...]): tuple of bookmarks identifying the rows.

        Returns:
            Tuple[int, ...]: A tuple of update counts containing one element for each row.
            The array is ordered according to the order in which bookmarks were given.
        """
        return self.__component.deleteRows(rows)  # type: ignore

    # endregion XDeleteRows
