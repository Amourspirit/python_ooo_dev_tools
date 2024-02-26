from __future__ import annotations
from typing import cast, TYPE_CHECKING
from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement
from ooodev.adapter.container.named_partial import NamedPartial
from ooodev.adapter.component_base import ComponentBase


if TYPE_CHECKING:
    from com.sun.star.text import TextFrame  # service
    from com.sun.star.text import XTextFrame


class TextFrameComp(ComponentBase, PropertyChangeImplement, VetoableChangeImplement, NamedPartial):
    """
    Class for managing TextFrame Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XTextFrame) -> None:
        """
        Constructor

        Args:
            component (TextFrame): UNO TextFrame Component that supports ``com.sun.star.text.TextFrame`` service.
        """

        ComponentBase.__init__(self, component)
        # pylint: disable=no-member
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        PropertyChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        VetoableChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        NamedPartial.__init__(self, component=self.component, interface=None)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.text.TextFrame",)

    # endregion Overrides

    # region Properties
    @property
    def component(self) -> TextFrame:
        """TextFrame Component"""
        # pylint: disable=no-member
        return cast("TextFrame", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
