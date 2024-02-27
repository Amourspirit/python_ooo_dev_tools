from __future__ import annotations
from typing import cast, TYPE_CHECKING
from ooodev.adapter.component_base import ComponentBase


if TYPE_CHECKING:
    from com.sun.star.drawing import TextProperties  # service


class TextPropertiesComp(ComponentBase):
    """
    Class for managing table TextProperties Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: TextProperties) -> None:
        """
        Constructor

        Args:
            component (TextProperties): UNO TextProperties Component.
        """
        ComponentBase.__init__(self, component)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.drawing.TextProperties",)

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> TextProperties:
        """TextProperties Component"""
        # pylint: disable=no-member
        return cast("TextProperties", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
