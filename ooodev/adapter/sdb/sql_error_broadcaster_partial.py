from __future__ import annotations
from typing import Any, TYPE_CHECKING

from com.sun.star.sdb import XSQLErrorBroadcaster

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.sdb import XSQLErrorListener
    from ooodev.utils.type_var import UnoInterface


class SQLErrorBroadcasterPartial:
    """
    Partial Class for ``XSQLErrorBroadcaster``.
    """

    def __init__(self, component: XSQLErrorBroadcaster, interface: UnoInterface | None = XSQLErrorBroadcaster) -> None:
        """
        Constructor

        Args:
            component (XSQLErrorBroadcaster): UNO Component that implements ``com.sun.star.sdb.XSQLErrorBroadcaster``.            interface (UnoInterface, optional): The interface to be validated. Defaults to ``None``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XSQLErrorBroadcaster
    def add_sql_error_listener(self, listener: XSQLErrorListener) -> None:
        """
        Adds the specified listener to receive the event \"errorOccurred\"
        """
        self.__component.addSQLErrorListener(listener)

    def remove_sql_error_listener(self, listener: XSQLErrorListener) -> None:
        """
        Removes the specified listener.
        """
        self.__component.removeSQLErrorListener(listener)

    # endregion XSQLErrorBroadcaster
