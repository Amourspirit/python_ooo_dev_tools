from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.awt import XDialog

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class DialogPartial:
    """
    Partial class for XDialog.
    """

    def __init__(self, component: XDialog, interface: UnoInterface | None = XDialog) -> None:
        """
        Constructor

        Args:
            component (XDialog): UNO Component that implements ``com.sun.star.awt.XDialog`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XDialog``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XDialog
    def end_execute(self) -> None:
        """
        Hides the dialog and then causes XDialog.execute() to return.
        """
        self.__component.endExecute()

    def execute(self) -> int:
        """
        Runs the dialog modally: shows it, and waits for the execution to end.

        Returns an exit code (e.g., indicating the button that was used to end the execution).
        """
        return self.__component.execute()

    def get_title(self) -> str:
        """
        Gets the title of the dialog.
        """
        return self.__component.getTitle()

    def set_title(self, title: str) -> None:
        """
        sets the title of the dialog.
        """
        self.__component.setTitle(title)

    # endregion XDialog
