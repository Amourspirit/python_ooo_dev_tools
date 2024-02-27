from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING

from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.container.container_partial import ContainerPartial
from ooodev.adapter.container.name_container_partial import NameContainerPartial
from ooodev.adapter.container.index_container_partial import IndexContainerPartial
from ooodev.adapter.container.enumeration_access_partial import EnumerationAccessPartial
from ooodev.adapter.script.event_attacher_manager_partial import EventAttacherManagerPartial

if TYPE_CHECKING:
    from com.sun.star.form import FormComponents


class FormComponentsComp(
    ComponentBase,
    ContainerPartial,
    NameContainerPartial,
    IndexContainerPartial,
    EnumerationAccessPartial,
    EventAttacherManagerPartial,
):
    """
    Class for managing FormComponents Service.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: Any) -> None:
        """
        Constructor

        Args:
            component (Any): UNO Component that implements ``com.sun.star.form.FormComponents`` service.
        """

        ComponentBase.__init__(self, component)
        ContainerPartial.__init__(self, component=self.component, interface=None)
        NameContainerPartial.__init__(self, component=self.component, interface=None)
        IndexContainerPartial.__init__(self, component=self.component, interface=None)
        EnumerationAccessPartial.__init__(self, component=self.component, interface=None)
        EventAttacherManagerPartial.__init__(self, component=self.component, interface=None)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.form.FormComponents",)

    # endregion Overrides

    # region Properties

    @property
    def component(self) -> FormComponents:
        """FormComponents Component"""
        # pylint: disable=no-member
        return cast("FormComponents", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
