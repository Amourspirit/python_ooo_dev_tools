from __future__ import annotations
from typing import cast, TYPE_CHECKING
from ooodev.adapter.awt.uno_control_time_field_comp import UnoControlTimeFieldComp
from ooodev.adapter.form.bound_control_partial import BoundControlPartial

if TYPE_CHECKING:
    from com.sun.star.form.control import TimeField


class TimeFieldComp(UnoControlTimeFieldComp, BoundControlPartial):
    """Class for TimeField Control"""

    def __init__(self, component: TimeField):
        """
        Constructor

        Args:
            component (Any): Component that implements ``com.sun.star.form.control.TimeField`` service.
        """
        UnoControlTimeFieldComp.__init__(self, component=component)
        BoundControlPartial.__init__(self, component=component, interface=None)

    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.form.control.TimeField",)

    @property
    def component(self) -> TimeField:
        """TimeField Component"""
        # pylint: disable=no-member
        return cast("TimeField", self._ComponentBase__get_component())  # type: ignore
