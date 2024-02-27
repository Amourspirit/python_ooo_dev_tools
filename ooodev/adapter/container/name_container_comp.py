from __future__ import annotations
from typing import cast, TYPE_CHECKING

from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.container.name_container_partial import NameContainerPartial

if TYPE_CHECKING:
    from com.sun.star.container import XNameContainer


class NameContainerComp(ComponentBase, NameContainerPartial):
    """
    Class for managing XNameContainer Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XNameContainer) -> None:
        """
        Constructor

        Args:
            component (XNameContainer): UNO Component that implements ``com.sun.star.container.XNameContainer``.
        """

        ComponentBase.__init__(self, component)
        NameContainerPartial.__init__(self, component=self.component)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ()

    # endregion Overrides

    # region Properties

    @property
    def component(self) -> XNameContainer:
        """XNameContainer Component"""
        # pylint: disable=no-member
        return cast("XNameContainer", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
