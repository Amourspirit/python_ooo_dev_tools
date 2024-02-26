from __future__ import annotations
from typing import cast, TYPE_CHECKING

from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.awt.uno_control_dialog_partial import UnoControlDialogPartial

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlDialog


class UnoControlDialogComp(ComponentBase, UnoControlDialogPartial):
    """
    Class for managing UnoControlDialog Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: UnoControlDialog) -> None:
        """
        Constructor

        Args:
            component (UnoControlDialog): UNO Component that implements ``com.sun.star.awt.UnoControlDialog`` service.
        """

        ComponentBase.__init__(self, component)
        UnoControlDialogPartial.__init__(self, component=self.component, interface=None)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.awt.UnoControlDialog",)

    # endregion Overrides

    # region Properties

    @property
    def component(self) -> UnoControlDialog:
        """UnoControlDialog Component"""
        return cast("UnoControlDialog", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
