from __future__ import annotations
from typing import cast, TYPE_CHECKING
import uno
from ooodev.adapter.awt.date_field_partial import DateFieldPartial
from ooodev.adapter.awt.uno_control_edit_comp import UnoControlEditComp

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlDateField


class UnoControlDateFieldComp(UnoControlEditComp, DateFieldPartial):

    def __init__(self, component: UnoControlDateField):
        """
        Constructor

        Args:
            component (Any): Component that implements ``com.sun.star.awt.UnoControlDateField`` service.
        """
        UnoControlEditComp.__init__(self, component=component)
        DateFieldPartial.__init__(self, component=self.component, interface=None)

    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.awt.UnoControlDateField",)

    @property
    def component(self) -> UnoControlDateField:
        """UnoControlDateField Component"""
        # pylint: disable=no-member
        return cast("UnoControlDateField", self._ComponentBase__get_component())  # type: ignore
