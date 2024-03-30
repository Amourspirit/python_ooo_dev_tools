from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
from ooodev.adapter.awt.uno_control_edit_comp import UnoControlEditComp
from ooodev.adapter.form.bound_control_partial import BoundControlPartial
from ooodev.adapter.form.change_events import ChangeEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs

if TYPE_CHECKING:
    from com.sun.star.form.control import TextField


class TextFieldComp(UnoControlEditComp, BoundControlPartial, ChangeEvents):
    """Class for TextField Control"""

    # pylint: disable=unused-argument

    def __init__(self, component: TextField):
        """
        Constructor

        Args:
            component (Any): Component that implements ``com.sun.star.form.control.TextField`` service.
        """
        # pylint: disable=no-member
        UnoControlEditComp.__init__(self, component=component)
        BoundControlPartial.__init__(self, component=component, interface=None)
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        ChangeEvents.__init__(self, trigger_args=generic_args, cb=self.__on_change_listener_add_remove)

    # region Lazy Listeners

    def __on_change_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.component.addChangeListener(self.events_listener_change)
        event.remove_callback = True

    # endregion Lazy Listeners

    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.form.control.TextField",)

    @property
    def component(self) -> TextField:
        """TextField Component"""
        # pylint: disable=no-member
        return cast("TextField", self._ComponentBase__get_component())  # type: ignore
