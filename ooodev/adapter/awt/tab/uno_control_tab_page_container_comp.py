from __future__ import annotations
from typing import cast, TYPE_CHECKING
import uno
from ooodev.adapter.awt.uno_control_comp import UnoControlComp
from ooodev.adapter.awt.tab.tab_page_container_partial import TabPageContainerPartial

if TYPE_CHECKING:
    from com.sun.star.awt.tab import UnoControlTabPageContainer


class UnoControlTabPageContainerComp(UnoControlComp, TabPageContainerPartial):

    def __init__(self, component: UnoControlTabPageContainer):
        """
        Constructor

        Args:
            component (Any): Component that implements ``com.sun.star.awt.UnoControlTabPageContainer`` service.
        """
        UnoControlComp.__init__(self, component=component)
        TabPageContainerPartial.__init__(self, component=self.component, interface=None)

    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.awt.tab.UnoControlTabPageContainer",)

    @property
    def component(self) -> UnoControlTabPageContainer:
        """UnoControlTabPageContainer Component"""
        # pylint: disable=no-member
        return cast("UnoControlTabPageContainer", self._ComponentBase__get_component())  # type: ignore
