from __future__ import annotations
from typing import cast, TYPE_CHECKING
from ooodev.adapter.style.style_comp import StyleComp
from ooodev.adapter.text.numbering_rules_comp import NumberingRulesComp
from ooodev.adapter.beans.properties_change_implement import PropertiesChangeImplement

if TYPE_CHECKING:
    from com.sun.star.text import NumberingStyle  # service
    from com.sun.star.style import XStyle


class NumberingStyleComp(NumberingRulesComp, StyleComp, PropertiesChangeImplement):
    """
    Class for managing table NumberingStyle Component.
    """

    def __init__(self, component: XStyle) -> None:
        """
        Constructor

        Args:
            component (XStyle): UNO Component that supports ``com.sun.star.style.NumberingStyle`` service.
        """
        StyleComp.__init__(self, component)
        NumberingRulesComp.__init__(self, component=component)  # type: ignore
        # pylint: disable=no-member
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        PropertiesChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.style.NumberingStyle", "com.sun.star.style.Style", "com.sun.star.text.NumberingRules")

    # endregion Overrides
    # region Properties
    if TYPE_CHECKING:

        @property
        def component(self) -> NumberingStyle:
            """NumberingStyle Component"""
            # pylint: disable=no-member
            return cast("NumberingStyle", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
