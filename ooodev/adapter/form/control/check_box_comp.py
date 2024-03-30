from __future__ import annotations
from typing import cast, TYPE_CHECKING
from ooodev.adapter.awt.uno_control_check_box_comp import UnoControlCheckBoxComp
from ooodev.adapter.form.bound_control_partial import BoundControlPartial

if TYPE_CHECKING:
    from com.sun.star.form.control import CheckBox


class CheckBoxComp(UnoControlCheckBoxComp, BoundControlPartial):
    """Class for CheckBox Control"""

    def __init__(self, component: CheckBox):
        """
        Constructor

        Args:
            component (Any): Component that implements ``com.sun.star.form.control.CheckBox`` service.
        """
        UnoControlCheckBoxComp.__init__(self, component=component)
        BoundControlPartial.__init__(self, component=component, interface=None)

    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.form.control.CheckBox",)

    @property
    def component(self) -> CheckBox:
        """CheckBox Component"""
        # pylint: disable=no-member
        return cast("CheckBox", self._ComponentBase__get_component())  # type: ignore
