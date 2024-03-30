from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
from ooodev.adapter.awt.uno_control_button_comp import UnoControlButtonComp
from ooodev.adapter.form.approve_action_events import ApproveActionEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs

if TYPE_CHECKING:
    from com.sun.star.form.control import CommandButton


class CommandButtonComp(UnoControlButtonComp, ApproveActionEvents):
    """Class for CommandButton Control"""

    # pylint: disable=unused-argument

    def __init__(self, component: CommandButton):
        """
        Constructor

        Args:
            component (Any): Component that implements ``com.sun.star.form.control.CommandButton`` service.
        """
        # pylint: disable=no-member
        UnoControlButtonComp.__init__(self, component=component)
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        ApproveActionEvents.__init__(self, trigger_args=generic_args, cb=self.__on_approve_action_add_remove)

    # region Lazy Listeners

    def __on_approve_action_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.component.addApproveActionListener(self.events_listener_approve_action)
        event.remove_callback = True

    # endregion Lazy Listeners

    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.form.control.CommandButton",)

    @property
    def component(self) -> CommandButton:
        """CommandButton Component"""
        # pylint: disable=no-member
        return cast("CommandButton", self._ComponentBase__get_component())  # type: ignore
