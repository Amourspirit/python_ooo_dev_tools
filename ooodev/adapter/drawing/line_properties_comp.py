from __future__ import annotations
from typing import cast, TYPE_CHECKING
from ooodev.adapter.component_base import ComponentBase

from ooodev.adapter.drawing.line_properties_partial import LinePropertiesPartial

if TYPE_CHECKING:
    from com.sun.star.drawing import LineProperties  # service


class LinePropertiesComp(ComponentBase, LinePropertiesPartial):
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
        LinePropertiesPartial.__init__(self, component=component)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        # Validated by LinePropertiesPartial
        return ()

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> LineProperties:
        """LineProperties Component"""
        # pylint: disable=no-member
        return cast("LineProperties", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
