from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno

from ooodev.adapter.container.index_container_comp import IndexContainerComp
from ooodev.adapter.lang.multi_service_factory_partial import MultiServiceFactoryPartial

if TYPE_CHECKING:
    from com.sun.star.ui import ActionTriggerContainer  # service

    # from com.sun.star.ui import ActionTrigger
    from ooodev.gui.menu.context.action_t import ActionT


class ActionTriggerContainerComp(IndexContainerComp["ActionT"], MultiServiceFactoryPartial):
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

    # region Methods
    def get_command_index(self, command_url: str) -> int:
        """
        Gets the index of the action trigger with the specified command URL.

        Args:
            command_url (str): Command URL

        Returns:
            int: Index of the action trigger with the specified command URL or ``-1`` if not found.
        """
        for i, itm in enumerate(self):
            if self.is_separator(itm):
                continue
            if itm.CommandURL == command_url:  # type: ignore
                return i
        return -1

    def is_separator(self, itm: ActionT) -> bool:
        """
        Determines if the specified action trigger is a separator.

        Args:
            itm (ActionT): Action Trigger

        Returns:
            bool: ``True`` if the specified action trigger is a separator, ``False`` otherwise.
        """
        return itm.supportsService("com.sun.star.ui.ActionTriggerSeparator")

    # endregion Methods

    # region Properties
    @property
    def component(self) -> ActionTriggerContainer:
        """ActionTriggerContainer Component"""
        # pylint: disable=no-member
        return cast("ActionTriggerContainer", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
