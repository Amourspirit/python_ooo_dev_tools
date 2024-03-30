from __future__ import annotations
from typing import TYPE_CHECKING
from ooodev.adapter.awt.uno_control_time_field_comp import UnoControlTimeFieldComp
from ooodev.adapter.awt.spin_field_partial import SpinFieldPartial

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlTimeField


class ViewTimeField(UnoControlTimeFieldComp, SpinFieldPartial):
    """View for the Time Field control."""

    def __init__(self, component: UnoControlTimeField):
        """
        Constructor

        Args:
            component (Any): Component that implements ``com.sun.star.awt.UnoControlTimeField`` service.
        """
        UnoControlTimeFieldComp.__init__(self, component=component)
        SpinFieldPartial.__init__(self, component=component)
