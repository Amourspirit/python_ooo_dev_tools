from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno

from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement
from ooodev.adapter.text.text_table_cursor_partial import TextTableCursorPartial
from ooodev.adapter.beans.property_set_partial import PropertySetPartial
from ooodev.adapter.style.character_properties_partial import CharacterPropertiesPartial
from ooodev.adapter.style.paragraph_properties_partial import ParagraphPropertiesPartial

if TYPE_CHECKING:
    from com.sun.star.text import TextTableCursor


class TextTableCursorComp(
    ComponentBase,
    TextTableCursorPartial,
    PropertySetPartial,
    CharacterPropertiesPartial,
    ParagraphPropertiesPartial,
    PropertyChangeImplement,
    VetoableChangeImplement,
):
    """
    Class for managing TextTableCursor Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: Any) -> None:
        """
        Constructor

        Args:
            component (Any): UNO Component that supports ``com.sun.star.text.TextTableCursor`` service.
        """

        ComponentBase.__init__(self, component)
        TextTableCursorPartial.__init__(self, component=component, interface=None)
        PropertySetPartial.__init__(self, component=component, interface=None)
        CharacterPropertiesPartial.__init__(self, component=component)
        ParagraphPropertiesPartial.__init__(self, component=component)
        # pylint: disable=no-member
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        PropertyChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        VetoableChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.text.TextTableCursor",)

    # endregion Overrides

    # region XEnumerationAccess

    # endregion XEnumerationAccess

    # region Properties
    @property
    def component(self) -> TextTableCursor:
        """TextTableCursor Component"""
        # pylint: disable=no-member
        return cast("TextTableCursor", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
