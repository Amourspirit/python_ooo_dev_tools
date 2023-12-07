from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement
from ooodev.adapter.beans.properties_change_implement import PropertiesChangeImplement
from ooodev.adapter.component_base import ComponentBase
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.adapter.container.named_partial import NamedPartial


if TYPE_CHECKING:
    from com.sun.star.style import CharacterStyle  # service
    from com.sun.star.style import XStyle


class CharacterStyleComp(
    ComponentBase, PropertyChangeImplement, VetoableChangeImplement, PropertiesChangeImplement, NamedPartial
):
    """
    Class for managing CharacterStyle Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XStyle) -> None:
        """
        Constructor

        Args:
            component (CharacterStyle): UNO Component that support ``com.sun.star.style.CharacterStyle`` service.
        """
        ComponentBase.__init__(self, component)
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        PropertyChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        VetoableChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        PropertiesChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        NamedPartial.__init__(self, component=self.component)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.style.CharacterStyle",)

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> CharacterStyle:
        """CharacterStyle Component"""
        return cast("CharacterStyle", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
