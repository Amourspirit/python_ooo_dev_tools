from __future__ import annotations
from typing import cast, TYPE_CHECKING
from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.container.enumeration_access_partial import EnumerationAccessPartial

if TYPE_CHECKING:
    from com.sun.star.frame import Components  # service


class ComponentsComp(ComponentBase, EnumerationAccessPartial):
    """
    Class for managing Components.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: Components) -> None:
        """
        Constructor

        Args:
            component (Components): UNO Component that implements ``com.sun.star.frame.Components`` service.
        """
        ComponentBase.__init__(self, component)
        EnumerationAccessPartial.__init__(self, component=component)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ()

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> Components:
        """Components Component"""
        # pylint: disable=no-member
        return cast("Components", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
