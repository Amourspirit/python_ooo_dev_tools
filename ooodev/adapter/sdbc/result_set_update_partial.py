from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.sdbc import XResultSetUpdate

from ooodev.exceptions import ex as mEx
from ooodev.utils import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class ResultSetUpdatePartial:
    """
    Partial class for XResultSetUpdate.
    """

    def __init__(self, component: XResultSetUpdate, interface: UnoInterface | None = XResultSetUpdate) -> None:
        """
        Constructor

        Args:
            component (XResultSetUpdate): UNO Component that implements ``com.sun.star.container.XResultSetUpdate`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XResultSetUpdate``.
        """
        self.__interface = interface
        self.__validate(component)
        self.__component = component

    def __validate(self, component: Any) -> None:
        """
        Validates the component.

        Args:
            component (Any): The component to be validated.
        """
        if self.__interface is None:
            return
        if not mLo.Lo.is_uno_interfaces(component, self.__interface):
            raise mEx.MissingInterfaceError(self.__interface)

    # region XResultSetUpdate
    def cancel_row_updates(self) -> None:
        """
        Cancels the updates made to a row.

        This method may be called after calling an updateXXX method(s) and before calling
        ``XResultSetUpdate.updateRow()`` to rollback the updates made to a row.
        If no updates have been made or updateRow has already been called, then this method has no effect.
        """
        self.__component.cancelRowUpdates()

    def delete_row(self) -> None:
        """
        Deletes the current row from the result set and the underlying database.
        """
        self.__component.deleteRow()

    def insert_row(self) -> None:
        """
        Inserts a new row into the result set and the underlying database.

        Must be on the insert row when this method is called.
        """
        self.__component.insertRow()

    def move_to_current_row(self) -> None:
        """
        Moves the cursor to the remembered cursor position, usually the current row.

        This method has no effect if the cursor is not on the insert row.
        """
        self.__component.moveToCurrentRow()

    def move_to_insert_row(self) -> None:
        """
        Moves the cursor to the insert row.

        The current row is remembered while the cursor is positioned on the insert row.
        """
        self.__component.moveToInsertRow()

    def update_row(self) -> None:
        """
        Updates the underlying database with the new contents of the current row.

        This method cannot be called when the cursor is on the insert row.
        """
        self.__component.updateRow()

    # endregion XResultSetUpdate
