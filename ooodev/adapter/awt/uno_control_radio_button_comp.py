from __future__ import annotations
from typing import cast, TYPE_CHECKING
import uno
from ooodev.adapter.awt.uno_control_comp import UnoControlComp
from ooodev.adapter.awt.radio_button_partial import RadioButtonPartial
from ooodev.adapter.awt.layout_constrains_partial import LayoutConstrainsPartial

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlRadioButton


class UnoControlRadioButtonComp(UnoControlComp, RadioButtonPartial, LayoutConstrainsPartial):

    def __init__(self, component: UnoControlRadioButton):
        """
        Constructor

        Args:
            component (Any): Component that implements ``com.sun.star.awt.UnoControlRadioButton`` service.
        """
        UnoControlComp.__init__(self, component=component)
        RadioButtonPartial.__init__(self, component=self.component, interface=None)
        LayoutConstrainsPartial.__init__(self, component=self.component, interface=None)

    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.awt.UnoControlRadioButton",)

    @property
    def component(self) -> UnoControlRadioButton:
        """UnoControlRadioButton Component"""
        # pylint: disable=no-member
        return cast("UnoControlRadioButton", self._ComponentBase__get_component())  # type: ignore
