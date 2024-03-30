from __future__ import annotations
from typing import cast, TYPE_CHECKING
import uno
from ooodev.adapter.awt.uno_control_comp import UnoControlComp
from ooodev.adapter.awt.check_box_partial import CheckBoxPartial
from ooodev.adapter.awt.layout_constrains_partial import LayoutConstrainsPartial

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlCheckBox


class UnoControlCheckBoxComp(UnoControlComp, CheckBoxPartial, LayoutConstrainsPartial):

    def __init__(self, component: UnoControlCheckBox):
        """
        Constructor

        Args:
            component (Any): Component that implements ``com.sun.star.awt.UnoControlCheckBox`` service.
        """
        UnoControlComp.__init__(self, component=component)
        CheckBoxPartial.__init__(self, component=self.component, interface=None)
        LayoutConstrainsPartial.__init__(self, component=self.component, interface=None)

    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.awt.UnoControlCheckBox",)

    @property
    def component(self) -> UnoControlCheckBox:
        """UnoControlCheckBox Component"""
        # pylint: disable=no-member
        return cast("UnoControlCheckBox", self._ComponentBase__get_component())  # type: ignore
