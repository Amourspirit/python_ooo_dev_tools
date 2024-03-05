from __future__ import annotations
from typing import cast, TYPE_CHECKING

from ooodev.adapter.style.style_comp import StyleComp
from ooodev.adapter.beans.properties_change_implement import PropertiesChangeImplement

if TYPE_CHECKING:
    from com.sun.star.style import CharacterStyle  # service
    from com.sun.star.style import XStyle


class CharacterStyleComp(StyleComp, PropertiesChangeImplement):
    """
    Class for managing CharacterStyle Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XStyle) -> None:
        """
        Constructor

        Args:
            component (XStyle): UNO Component that support ``com.sun.star.style.CharacterStyle`` service.
        """
        StyleComp.__init__(self, component)
        # pylint: disable=no-member
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        PropertiesChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.style.CharacterStyle",)

    # endregion Overrides
    # region Properties
    if TYPE_CHECKING:

        @property
        def component(self) -> CharacterStyle:
            """CharacterStyle Component"""
            # pylint: disable=no-member
            return cast("CharacterStyle", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
