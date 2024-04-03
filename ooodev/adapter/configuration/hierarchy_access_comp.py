from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING

from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.container.name_access_partial import NameAccessPartial
from ooodev.adapter.container.hierarchical_name_access_partial import HierarchicalNameAccessPartial
from ooodev.adapter.container.container_partial import ContainerPartial
from ooodev.adapter.beans.exact_name_partial import ExactNamePartial

if TYPE_CHECKING:
    from com.sun.star.configuration import HierarchyAccess  # service


class HierarchyAccessComp(
    ComponentBase, NameAccessPartial, HierarchicalNameAccessPartial, ContainerPartial, ExactNamePartial
):
    """
    Class for managing HierarchyAccess Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: Any) -> None:
        """
        Constructor

        Args:
            component (XNameAccess): UNO Component that implements ``com.sun.star.container.XNameAccess``.
        """

        ComponentBase.__init__(self, component)
        NameAccessPartial.__init__(self, component=component, interface=None)
        HierarchicalNameAccessPartial.__init__(self, component=component, interface=None)
        ContainerPartial.__init__(self, component=component, interface=None)
        ExactNamePartial.__init__(self, component=component, interface=None)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.configuration.HierarchyAccess",)

    # endregion Overrides

    # region Properties

    @property
    def component(self) -> HierarchyAccess:
        """HierarchyAccess Component"""
        # pylint: disable=no-member
        return cast("HierarchyAccess", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties


def get_builder(component: Any, lo_inst: Any = None) -> Any:
    # pylint: disable=import-outside-toplevel
    from ooodev.utils.builder.default_builder import DefaultBuilder

    builder = DefaultBuilder(component, lo_inst)
    builder.add_import(
        name="ooodev.adapter.configuration.hierarchy_access_comp.HierarchyAccessComp",
        uno_name="com.sun.star.configuration.HierarchyAccess",
        optional=False,
        init_kind=1,
    )
    builder.add_import(
        name="ooodev.adapter.beans.property_set_info_partial.PropertySetInfoPartial",
        uno_name="com.sun.star.beans.XPropertySetInfo",
        optional=True,
    )
    builder.add_import(
        name="ooodev.adapter.beans.property_state_partial.PropertyStatePartial",
        uno_name="com.sun.star.beans.XPropertyState",
        optional=True,
    )
    builder.add_import(
        name="ooodev.adapter.beans.multi_property_states_partial.MultiPropertyStatesPartial",
        uno_name="com.sun.star.beans.XMultiPropertyStates",
        optional=True,
    )
    return builder
