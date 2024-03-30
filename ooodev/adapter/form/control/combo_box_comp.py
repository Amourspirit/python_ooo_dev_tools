from __future__ import annotations
from typing import cast, TYPE_CHECKING
from ooodev.adapter.awt.uno_control_combo_box_comp import UnoControlComboBoxComp
from ooodev.adapter.form.bound_control_partial import BoundControlPartial

if TYPE_CHECKING:
    from com.sun.star.form.control import ComboBox


class ComboBoxComp(UnoControlComboBoxComp, BoundControlPartial):
    """Class for ComboBox Control"""

    def __init__(self, component: ComboBox):
        """
        Constructor

        Args:
            component (Any): Component that implements ``com.sun.star.form.control.ComboBox`` service.
        """
        UnoControlComboBoxComp.__init__(self, component=component)
        BoundControlPartial.__init__(self, component=component, interface=None)

    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.form.control.ComboBox",)

    @property
    def component(self) -> ComboBox:
        """ComboBox Component"""
        # pylint: disable=no-member
        return cast("ComboBox", self._ComponentBase__get_component())  # type: ignore
