from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.adapter.lang.event_events import EventEvents
from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement
from ooodev.adapter.beans.property_set_partial import PropertySetPartial
from ooodev.adapter.drawing.shape_comp import ShapeComp
from ooodev.adapter.drawing.control_shape_partial import ControlShapePartial


if TYPE_CHECKING:
    from com.sun.star.drawing import ControlShape  # service


class ControlShapeComp(
    ShapeComp, ControlShapePartial, PropertySetPartial, EventEvents, PropertyChangeImplement, VetoableChangeImplement
):
    """
    Class for managing Shape which contains a control.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: ControlShape) -> None:
        """
        Constructor

        Args:
            component (ControlShape): UNO Volatile Result Component
        """
        ShapeComp.__init__(self, component)
        ControlShapePartial.__init__(self, component=component, interface=None)
        PropertySetPartial.__init__(self, component=self.component, interface=None)
        # pylint: disable=no-member
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        EventEvents.__init__(self, trigger_args=generic_args, cb=self._on_event_add_remove)
        PropertyChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        VetoableChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)

    # region Lazy Listeners
    def _on_event_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.component.addEventListener(self.events_listener_event)
        event.remove_callback = True

    # endregion Lazy Listeners

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.sheet.ControlShape",)

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> ControlShape:
        """Control Shape Component"""
        # pylint: disable=no-member
        return cast("ControlShape", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
