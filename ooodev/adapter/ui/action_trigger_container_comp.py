from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno

from ooodev.adapter.container.index_container_comp import IndexContainerComp
from ooodev.adapter.lang.multi_service_factory_partial import MultiServiceFactoryPartial

if TYPE_CHECKING:
    from com.sun.star.ui import ActionTriggerContainer  # service
    from com.sun.star.ui import ActionTrigger


class ActionTriggerContainerComp(IndexContainerComp["ActionTrigger"], MultiServiceFactoryPartial):
    """
    Class for managing Action Trigger Container Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: Any) -> None:
        """
        Constructor

        Args:
            component (Any): UNO Component that supports ``com.sun.star.ui.ActionTriggerContainer`` service.
        """

        IndexContainerComp.__init__(self, component)
        MultiServiceFactoryPartial.__init__(self, component, interface=None)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.ui.ActionTriggerContainer",)

    # endregion Overrides

    # region XEnumerationAccess

    # endregion XEnumerationAccess

    # region Properties
    @property
    def component(self) -> ActionTriggerContainer:
        """ActionTriggerContainer Component"""
        # pylint: disable=no-member
        return cast("ActionTriggerContainer", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
