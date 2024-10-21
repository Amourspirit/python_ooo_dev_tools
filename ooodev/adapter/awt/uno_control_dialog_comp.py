from __future__ import annotations
from typing import cast, TYPE_CHECKING

try:
    # python 3.12+
    from typing import override  # noqa # type: ignore
except ImportError:
    from typing_extensions import override  # noqa # type: ignore

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
    @override
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.awt.UnoControlDialog",)

    # endregion Overrides

    # region Properties

    @property
    @override
    def component(self) -> UnoControlDialog:
        """UnoControlDialog Component"""
        # pylint: disable=no-member
        return cast("UnoControlDialog", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
