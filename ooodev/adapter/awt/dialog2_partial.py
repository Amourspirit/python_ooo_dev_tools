from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.awt import XDialog2
from ooodev.adapter.awt.dialog_partial import DialogPartial


if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface
else:
    UnoInterface = Any


class Dialog2Partial(DialogPartial):
    """
    Partial Class for XDialog2.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XDialog2, interface: UnoInterface | None = XDialog2) -> None:
        """
        Constructor

        Args:
            component (XDialog2): UNO Component that implements ``com.sun.star.awt.XDialog2`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XDialog2``.
        """
        DialogPartial.__init__(self, component=component, interface=interface)
        self.__component = component

    # region XDialog2
    def end_dialog(self, result: int) -> None:
        """
        Hides the dialog and then causes XDialog.execute() to return with the given result value.
        """
        self.__component.endDialog(result)

    def set_help_id(self, help_id: str) -> None:
        """
        Sets the help id so that the standard help button action will show the appropriate help page.
        """
        self.__component.setHelpId(help_id)

    # endregion XDialog2
