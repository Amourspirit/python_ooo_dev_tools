from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.text.text_partial import TextPartial


if TYPE_CHECKING:
    from com.sun.star.drawing import Text  # service


class TextComp(ComponentBase, TextPartial):
    """
    Class for managing table Text Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: Any) -> None:
        """
        Constructor

        Args:
            component (Any): UNO component that supports ``com.sun.star.drawing.Text`` service.
        """
        ComponentBase.__init__(self, component)
        TextPartial.__init__(self, component, interface=None)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.drawing.Text",)

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> Text:
        """Text Component"""
        # pylint: disable=no-member
        return cast("Text", self._ComponentBase__get_component())  # type: ignore

        # endregion Properties
