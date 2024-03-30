from __future__ import annotations
from typing import TYPE_CHECKING
from ooodev.adapter.awt.uno_control_date_field_comp import UnoControlDateFieldComp
from ooodev.adapter.awt.spin_field_partial import SpinFieldPartial

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlDateField


class ViewDateField(UnoControlDateFieldComp, SpinFieldPartial):
    """View for the Date Field control."""

    def __init__(self, component: UnoControlDateField):
        """
        Constructor

        Args:
            component (Any): Component that implements ``com.sun.star.awt.UnoControlDateField`` service.
        """
        UnoControlDateFieldComp.__init__(self, component=component)
        SpinFieldPartial.__init__(self, component=component)
