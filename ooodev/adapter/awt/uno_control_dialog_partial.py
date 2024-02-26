from __future__ import annotations
from typing import TYPE_CHECKING
import uno

from com.sun.star.awt import XUnoControlDialog
from ooodev.adapter.container.control_container_partial import ControlContainerPartial
from ooodev.adapter.awt.dialog2_partial import Dialog2Partial
from ooodev.adapter.awt.control_partial import ControlPartial
from ooodev.adapter.awt.window_partial import WindowPartial
from ooodev.adapter.awt.top_window_partial import TopWindowPartial


if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class UnoControlDialogPartial(
    ControlContainerPartial,
    Dialog2Partial,
    WindowPartial,
    ControlPartial,
    TopWindowPartial,
):
    """
    Partial Class for XUnoControlDialog.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XUnoControlDialog, interface: UnoInterface | None = XUnoControlDialog) -> None:
        """
        Constructor

        Args:
            component (XUnoControlDialog): UNO Component that implements ``com.sun.star.awt.XUnoControlDialog`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XUnoControlDialog``.
        """
        ControlContainerPartial.__init__(self, component=component, interface=interface)
        Dialog2Partial.__init__(self, component=component, interface=interface)
        ControlPartial.__init__(self, component=component, interface=interface)
        WindowPartial.__init__(self, component=component, interface=interface)
        TopWindowPartial.__init__(self, component=component, interface=interface)
        self.__component = component

    # region XUnoControlDialog
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

    # endregion XUnoControlDialog
