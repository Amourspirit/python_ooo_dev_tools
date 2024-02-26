from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
from ooodev.adapter.awt.tab_controller_model_partial import TabControllerModelPartial
from ooodev.adapter.beans.property_bag_partial import PropertyBagPartial
from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement
from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.container.container_events import ContainerEvents
from ooodev.adapter.container.enumeration_access_partial import EnumerationAccessPartial
from ooodev.adapter.container.index_container_partial import IndexContainerPartial
from ooodev.adapter.container.name_container_partial import NameContainerPartial
from ooodev.adapter.container.named_partial import NamedPartial
from ooodev.adapter.form.form_partial import FormPartial
from ooodev.adapter.io.persist_object_partial import PersistObjectPartial
from ooodev.adapter.lang.component_partial import ComponentPartial
from ooodev.adapter.lang.event_events import EventEvents
from ooodev.adapter.script.event_attacher_manager_partial import EventAttacherManagerPartial
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.adapter.container.element_index_partial import ElementIndexPartial

if TYPE_CHECKING:
    from com.sun.star.form.component import Form  # service


class FormComp(
    ComponentBase,
    ComponentPartial,
    NamedPartial,
    PersistObjectPartial,
    PropertyBagPartial,
    NameContainerPartial,
    IndexContainerPartial,
    EnumerationAccessPartial,
    EventAttacherManagerPartial,
    TabControllerModelPartial,
    FormPartial,
    EventEvents,
    ContainerEvents,
    PropertyChangeImplement,
    VetoableChangeImplement,
    ElementIndexPartial,
):
    """
    Class for managing Form Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: Form) -> None:
        """
        Constructor

        Args:
            component (SheetCell): UNO Component that supports ``com.sun.star.form.component.Form`` service.
        """
        ComponentBase.__init__(self, component)
        ComponentPartial.__init__(self, component=component, interface=None)
        NamedPartial.__init__(self, component=component, interface=None)
        PersistObjectPartial.__init__(self, component=component, interface=None)
        PropertyBagPartial.__init__(self, component=component, interface=None)
        NameContainerPartial.__init__(self, component=component, interface=None)
        IndexContainerPartial.__init__(self, component=component, interface=None)
        EnumerationAccessPartial.__init__(self, component=component, interface=None)
        EventAttacherManagerPartial.__init__(self, component=component, interface=None)
        TabControllerModelPartial.__init__(self, component=component, interface=None)
        FormPartial.__init__(self, component=component, interface=None)
        # pylint: disable=no-member
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        EventEvents.__init__(self, trigger_args=generic_args, cb=self._on_event_events_add_remove)
        ContainerEvents.__init__(self, trigger_args=generic_args, cb=self._on_container_events_add_remove)
        PropertyChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        VetoableChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        ElementIndexPartial.__init__(self, component=self)

    # region Lazy Listeners

    def _on_container_events_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.component.addContainerListener(self.events_listener_container)
        event.remove_callback = True

    def _on_event_events_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.component.addEventListener(self.events_listener_event)
        event.remove_callback = True

    # endregion Lazy Listeners

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.form.component.Form",)

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> Form:
        """Form Component"""
        # pylint: disable=no-member
        return cast("Form", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
