from __future__ import annotations
from typing import cast, TYPE_CHECKING

from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.container.index_container_partial import IndexContainerPartial

if TYPE_CHECKING:
    from com.sun.star.container import XIndexContainer


class IndexContainerComp(ComponentBase, IndexContainerPartial):
    """
    Class for managing XIndexContainer Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XIndexContainer) -> None:
        """
        Constructor

        Args:
            component (XIndexContainer): UNO Component that implements ``com.sun.star.container.XIndexContainer``.
        """

        ComponentBase.__init__(self, component)
        IndexContainerPartial.__init__(self, component=self.component)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ()

    # endregion Overrides

    # region Properties

    @property
    def component(self) -> XIndexContainer:
        """XIndexContainer Component"""
        # pylint: disable=no-member
        return cast("XIndexContainer", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
