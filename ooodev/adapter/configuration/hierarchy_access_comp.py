from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING

from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.container import name_access_partial
from ooodev.adapter.container import hierarchical_name_access_partial
from ooodev.adapter.container import container_partial
from ooodev.adapter.beans import exact_name_partial

if TYPE_CHECKING:
    from com.sun.star.configuration import HierarchyAccess  # service


class HierarchyAccessComp(
    ComponentBase,
    name_access_partial.NameAccessPartial,
    hierarchical_name_access_partial.HierarchicalNameAccessPartial,
    container_partial.ContainerPartial,
    exact_name_partial.ExactNamePartial,
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
        name_access_partial.NameAccessPartial.__init__(self, component=component, interface=None)
        hierarchical_name_access_partial.HierarchicalNameAccessPartial.__init__(
            self, component=component, interface=None
        )
        container_partial.ContainerPartial.__init__(self, component=component, interface=None)
        exact_name_partial.ExactNamePartial.__init__(self, component=component, interface=None)

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


def get_builder(component: Any, lo_inst: Any = None, **kwargs) -> Any:
    # pylint: disable=import-outside-toplevel
    from ooodev.utils.builder.default_builder import DefaultBuilder

    builder = DefaultBuilder(component, lo_inst)
    local = kwargs.get("local", False)

    # region Excludes
    na_ex = cast(DefaultBuilder, name_access_partial.get_builder(component, lo_inst))
    ha_ex = cast(DefaultBuilder, hierarchical_name_access_partial.get_builder(component, lo_inst))
    container_ex = cast(DefaultBuilder, container_partial.get_builder(component, lo_inst))
    exact_ex = cast(DefaultBuilder, exact_name_partial.get_builder(component, lo_inst))
    builder.omits.update(na_ex.omits)
    builder.omits.update(ha_ex.omits)
    builder.omits.update(container_ex.omits)
    builder.omits.update(exact_ex.omits)

    if local:
        builder.set_omit(*na_ex.get_import_names())
        builder.set_omit(*ha_ex.get_import_names())
        builder.set_omit(*container_ex.get_import_names())
        builder.set_omit(*exact_ex.get_import_names())
    else:
        builder.add_from_instance(na_ex, make_optional=True)
        builder.add_from_instance(ha_ex, make_optional=True)
        builder.add_from_instance(container_ex, make_optional=True)
        builder.add_from_instance(exact_ex, make_optional=True)
    # endregion Excludes

    # in a from_lo method in this class the HierarchyAccessComp would be removed and used as the base class

    # builder.add_import(
    #     name="ooodev.adapter.configuration.hierarchy_access_comp.HierarchyAccessComp",
    #     uno_name="com.sun.star.configuration.HierarchyAccess",
    #     optional=False,
    #     init_kind=1,
    # )
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
    builder.add_event(
        module_name="ooodev.adapter.container.container_events",
        class_name="ContainerEvents",
        uno_name="com.sun.star.container.XContainer",
        optional=True,
    )
    return builder
