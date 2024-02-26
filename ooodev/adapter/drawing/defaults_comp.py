from __future__ import annotations
from typing import cast, TYPE_CHECKING
from ooodev.adapter.component_base import ComponentBase


if TYPE_CHECKING:
    from com.sun.star.text import Defaults  # service


class DefaultsComp(ComponentBase):
    """
    Class for managing table Defaults Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: Defaults) -> None:
        """
        Constructor

        Args:
            component (Defaults): UNO Defaults Component.
        """
        ComponentBase.__init__(self, component)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.drawing.Defaults",)

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> Defaults:
        """Defaults Component"""
        # pylint: disable=no-member
        return cast("Defaults", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
