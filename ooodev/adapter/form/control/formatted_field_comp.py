from __future__ import annotations
from typing import cast, TYPE_CHECKING
from ooodev.adapter.awt.uno_control_formatted_field_comp import UnoControlFormattedFieldComp
from ooodev.adapter.form.bound_control_partial import BoundControlPartial

if TYPE_CHECKING:
    from com.sun.star.form.control import FormattedField


class FormattedFieldComp(UnoControlFormattedFieldComp, BoundControlPartial):
    """Class for Formatted Field Control"""

    def __init__(self, component: FormattedField):
        """
        Constructor

        Args:
            component (Any): Component that implements ``com.sun.star.form.control.FormattedField`` service.
        """
        UnoControlFormattedFieldComp.__init__(self, component=component)
        BoundControlPartial.__init__(self, component=component, interface=None)

    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.form.control.FormattedField",)

    @property
    def component(self) -> FormattedField:
        """FormattedField Component"""
        # pylint: disable=no-member
        return cast("FormattedField", self._ComponentBase__get_component())  # type: ignore
