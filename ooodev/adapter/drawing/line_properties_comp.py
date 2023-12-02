from __future__ import annotations
from typing import cast, TYPE_CHECKING
from ooodev.adapter.component_base import ComponentBase


if TYPE_CHECKING:
    from com.sun.star.drawing import LineProperties  # service


class LinePropertiesComp(ComponentBase):
    """
    Class for managing table LineProperties Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: LineProperties) -> None:
        """
        Constructor

        Args:
            component (LineProperties): UNO LineProperties Component.
        """
        ComponentBase.__init__(self, component)

    # region Overrides
    def _get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.drawing.LineProperties",)

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> LineProperties:
        """LineProperties Component"""
        return cast("LineProperties", self._get_component())

    # endregion Properties
