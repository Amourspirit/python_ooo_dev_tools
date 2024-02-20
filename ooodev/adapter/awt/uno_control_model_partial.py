from __future__ import annotations
from typing import Any

from ooodev.adapter.beans.property_set_partial import PropertySetPartial
from ooodev.utils.partial.model_prop_partial import ModelPropPartial


class UnoControlModelPartial(PropertySetPartial):
    """Partial class for UnoControlModel."""

    def __init__(self):
        """
        Constructor

        Args:
            component (Any): Component that implements ``com.sun.star.awt.UnoControlModel`` service.
        """
        if not isinstance(self, ModelPropPartial):
            raise TypeError("This class must be used as a mixin that implements ModelPropPartial.")
        # pylint: disable=no-member
        PropertySetPartial.__init__(self, component=self.model, interface=None)
