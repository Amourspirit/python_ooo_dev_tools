from __future__ import annotations
from typing import cast, TYPE_CHECKING
from ooodev.adapter.component_base import ComponentBase


if TYPE_CHECKING:
    from com.sun.star.drawing import FillProperties  # service


class FillPropertiesComp(ComponentBase):
    """
    Class for managing table FillProperties Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: FillProperties) -> None:
        """
        Constructor

        Args:
            component (FillProperties): UNO FillProperties Component.
        """
        ComponentBase.__init__(self, component)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.drawing.FillProperties",)

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> FillProperties:
        """FillProperties Component"""
        return cast("FillProperties", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
