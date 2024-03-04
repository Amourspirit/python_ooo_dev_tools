from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.sdb import XResultSetAccess

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.sdbc import XResultSet
    from ooodev.utils.type_var import UnoInterface


class ResultSetAccessPartial:
    """
    Partial class for XResultSetAccess.
    """

    def __init__(self, component: XResultSetAccess, interface: UnoInterface | None = XResultSetAccess) -> None:
        """
        Constructor

        Args:
            component (XResultSetAccess): UNO Component that implements ``com.sun.star.container.XResultSetAccess`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XResultSetAccess``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XResultSetAccess
    def create_result_set(self) -> XResultSet:
        """
        Returns a new ResultSet based on the object.

        Returns:
            XResultSet: The new result set.
        """
        return self.__component.createResultSet()

    # endregion XResultSetAccess
