from __future__ import annotations
from typing import cast, TYPE_CHECKING
import uno
from ooodev.adapter.awt.uno_control_comp import UnoControlComp
from ooodev.adapter.awt.uno_control_container_partial import UnoControlContainerPartial
from ooodev.adapter.awt.control_container_partial import ControlContainerPartial
from ooodev.adapter.container.container_partial import ContainerPartial

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlContainer


class UnoControlContainerComp(UnoControlComp, UnoControlContainerPartial, ControlContainerPartial, ContainerPartial):

    def __init__(self, component: UnoControlContainer):
        """
        Constructor

        Args:
            component (Any): Component that implements ``com.sun.star.awt.UnoControlContainer`` service.
        """
        UnoControlComp.__init__(self, component=component)
        UnoControlContainerPartial.__init__(self, component=self.component, interface=None)
        ControlContainerPartial.__init__(self, component=self.component, interface=None)
        ContainerPartial.__init__(self, component=self.component, interface=None)

    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.awt.UnoControlContainer",)

    @property
    def component(self) -> UnoControlContainer:
        """UnoControlContainer Component"""
        # pylint: disable=no-member
        return cast("UnoControlContainer", self._ComponentBase__get_component())  # type: ignore
