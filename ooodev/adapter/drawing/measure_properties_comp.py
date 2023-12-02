from __future__ import annotations
from typing import cast, TYPE_CHECKING
from ooodev.adapter.component_base import ComponentBase


if TYPE_CHECKING:
    from com.sun.star.drawing import MeasureProperties  # service


class MeasurePropertiesComp(ComponentBase):
    """
    Class for managing table MeasureProperties Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: MeasureProperties) -> None:
        """
        Constructor

        Args:
            component (MeasureProperties): UNO MeasureProperties Component.
        """
        ComponentBase.__init__(self, component)

    # region Overrides
    def _get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.drawing.MeasureProperties",)

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> MeasureProperties:
        """MeasureProperties Component"""
        return cast("MeasureProperties", self._get_component())

    # endregion Properties
