from __future__ import annotations
from typing import cast, TYPE_CHECKING
import uno
from ooodev.adapter.awt.uno_control_comp import UnoControlComp
from ooodev.adapter.awt.progress_bar_partial import ProgressBarPartial

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlProgressBar


class UnoControlProgressBarComp(UnoControlComp, ProgressBarPartial):

    def __init__(self, component: UnoControlProgressBar):
        """
        Constructor

        Args:
            component (Any): Component that implements ``com.sun.star.awt.UnoControlProgressBar`` service.
        """
        UnoControlComp.__init__(self, component=component)
        ProgressBarPartial.__init__(self, component=self.component, interface=None)

    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.awt.UnoControlProgressBar",)

    @property
    def component(self) -> UnoControlProgressBar:
        """UnoControlProgressBar Component"""
        # pylint: disable=no-member
        return cast("UnoControlProgressBar", self._ComponentBase__get_component())  # type: ignore
