from __future__ import annotations
from typing import cast, TYPE_CHECKING
import uno
from ooodev.adapter.awt.uno_control_comp import UnoControlComp
from ooodev.adapter.awt.scroll_bar_partial import ScrollBarPartial

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlScrollBar


class UnoControlScrollBarComp(UnoControlComp, ScrollBarPartial):

    def __init__(self, component: UnoControlScrollBar):
        """
        Constructor

        Args:
            component (Any): Component that implements ``com.sun.star.awt.UnoControlScrollBar`` service.
        """
        UnoControlComp.__init__(self, component=component)
        ScrollBarPartial.__init__(self, component=self.component, interface=None)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.awt.UnoControlScrollBar",)

    # endregion Overrides

    @property
    def component(self) -> UnoControlScrollBar:
        """UnoControlScrollBar Component"""
        # pylint: disable=no-member
        return cast("UnoControlScrollBar", self._ComponentBase__get_component())  # type: ignore
