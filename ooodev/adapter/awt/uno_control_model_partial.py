from __future__ import annotations
from typing import TYPE_CHECKING
import uno

from ooodev.adapter.beans.property_set_partial import PropertySetPartial
from ooodev.utils.partial.model_prop_partial import ModelPropPartial
from ooodev.adapter.awt.control_model_partial import ControlModelPartial
from ooodev.adapter.util.cloneable_partial import CloneablePartial

if TYPE_CHECKING:
    from com.sun.star.lang import XEventListener


class UnoControlModelPartial(ControlModelPartial, PropertySetPartial, CloneablePartial):
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
        ControlModelPartial.__init__(self)
        PropertySetPartial.__init__(self, component=self.model, interface=None)
        CloneablePartial.__init__(self, component=self.model, interface=None)

    # region XComponent
    def add_event_listener(self, listener: XEventListener) -> None:
        """
        Adds an event listener to the component.

        Args:
            listener (XEventListener): The event listener to be added.
        """
        # pylint: disable=no-member
        self.model.addEventListener(listener)  # type: ignore

    def remove_event_listener(self, listener: XEventListener) -> None:
        """
        Removes an event listener from the component.

        Args:
            listener (XEventListener): The event listener to be removed.
        """
        # pylint: disable=no-member
        self.model.removeEventListener(listener)  # type: ignore

    def dispose(self) -> None:
        """
        Disposes the component.
        """
        # pylint: disable=no-member
        self.model.dispose()  # type: ignore

    # endregion XComponent
