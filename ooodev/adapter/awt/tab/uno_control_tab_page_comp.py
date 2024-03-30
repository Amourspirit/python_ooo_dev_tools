from __future__ import annotations
from typing import cast, TYPE_CHECKING
import uno
from ooodev.adapter.awt.uno_control_container_comp import UnoControlContainerComp
from ooodev.adapter.awt.tab.tab_page_partial import TabPagePartial

if TYPE_CHECKING:
    from com.sun.star.awt.tab import UnoControlTabPage


class UnoControlTabPageComp(UnoControlContainerComp, TabPagePartial):

    def __init__(self, component: UnoControlTabPage):
        """
        Constructor

        Args:
            component (Any): Component that implements ``com.sun.star.awt.UnoControlTabPage`` service.
        """
        # TabPagePartial is an empty class because XTabPage has no methods.
        UnoControlContainerComp.__init__(self, component=component)
        TabPagePartial.__init__(self)

    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.awt.tab.UnoControlTabPage",)

    @property
    def component(self) -> UnoControlTabPage:
        """UnoControlTabPage Component"""
        # pylint: disable=no-member
        return cast("UnoControlTabPage", self._ComponentBase__get_component())  # type: ignore
