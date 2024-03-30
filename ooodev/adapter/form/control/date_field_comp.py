from __future__ import annotations
from typing import cast, TYPE_CHECKING
from ooodev.adapter.awt.uno_control_date_field_comp import UnoControlDateFieldComp
from ooodev.adapter.form.bound_control_partial import BoundControlPartial

if TYPE_CHECKING:
    from com.sun.star.form.control import DateField


class DateFieldComp(UnoControlDateFieldComp, BoundControlPartial):
    """Class for DateField Control"""

    def __init__(self, component: DateField):
        """
        Constructor

        Args:
            component (Any): Component that implements ``com.sun.star.form.control.DateField`` service.
        """
        UnoControlDateFieldComp.__init__(self, component=component)
        BoundControlPartial.__init__(self, component=component, interface=None)

    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.form.control.DateField",)

    @property
    def component(self) -> DateField:
        """DateField Component"""
        # pylint: disable=no-member
        return cast("DateField", self._ComponentBase__get_component())  # type: ignore
