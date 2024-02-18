from __future__ import annotations
from typing import Any

from ooodev.adapter.beans.property_set_partial import PropertySetPartial


class UnoControlModelPartial(PropertySetPartial):
    """Partial class for UnoControlModel."""

    def __init__(self, component: Any):
        """
        Constructor

        Args:
            component (Any): Component that implements ``com.sun.star.awt.UnoControlModel`` service.
        """
        PropertySetPartial.__init__(self, component=component, interface=None)
