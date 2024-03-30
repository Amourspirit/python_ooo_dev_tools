from __future__ import annotations
from typing import TYPE_CHECKING
from ooodev.adapter.awt.uno_control_formatted_field_comp import UnoControlFormattedFieldComp
from ooodev.adapter.awt.spin_field_partial import SpinFieldPartial

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlFormattedField


class ViewFormattedField(UnoControlFormattedFieldComp, SpinFieldPartial):
    """View for the Formatted Field control."""

    def __init__(self, component: UnoControlFormattedField):
        """
        Constructor

        Args:
            component (Any): Component that implements ``com.sun.star.awt.UnoControlFormattedField`` service.
        """
        UnoControlFormattedFieldComp.__init__(self, component=component)
        SpinFieldPartial.__init__(self, component=component)
