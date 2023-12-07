from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement
from ooodev.adapter.component_base import ComponentBase


if TYPE_CHECKING:
    from com.sun.star.text import TextRange  # service
    from com.sun.star.text import XTextRange


class TextRangeComp(ComponentBase, PropertyChangeImplement, VetoableChangeImplement):
    """
    Class for managing TextRange Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XTextRange) -> None:
        """
        Constructor

        Args:
            component (TextRange): UNO TextRange Component that supports ``com.sun.star.text.TextRange`` service.
        """

        ComponentBase.__init__(self, component)
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        PropertyChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        VetoableChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.text.TextRange",)

    # endregion Overrides

    # region Methods
    def get_start(self) -> TextRangeComp:
        """Returns a text range which contains only the start of this text range."""
        return TextRangeComp(self.component.getStart())

    def get_end(self) -> TextRangeComp:
        """Returns a text range which contains only the end of this text range."""
        return TextRangeComp(self.component.getEnd())

    def get_string(self) -> str:
        """Returns the string of this text range."""
        return self.component.getString()

    def set_string(self, string: str) -> None:
        """
        Sets the string of this text range.

        The whole string of characters of this piece of text is replaced.
        All styles are removed when applying this method.
        """
        self.component.setString(string)

    # endregion Methods

    # region Properties
    @property
    def component(self) -> TextRange:
        """TextRange Component"""
        return cast("TextRange", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
