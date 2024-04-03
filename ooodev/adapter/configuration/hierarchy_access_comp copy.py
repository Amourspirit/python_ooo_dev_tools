from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING

from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.container.hierarchical_name_access_partial import HierarchicalNameAccessPartial
from ooodev.adapter.beans import exact_name_partial
from ooodev.adapter.container import name_access_partial
from ooodev.adapter.container import hierarchical_name_access_partial
from ooodev.adapter.container import hierarchical_name_partial
from ooodev.adapter.container import named_partial
from ooodev.adapter.container import container_partial
from ooodev.adapter.beans import property_partial
from ooodev.adapter.beans import property_set_info_partial
from ooodev.adapter.beans import property_state_partial
from ooodev.adapter.beans import multi_property_states_partial
from ooodev.adapter.beans import property_with_state_partial
from ooodev.adapter.container import child_partial

if TYPE_CHECKING:
    from com.sun.star.configuration import ConfigurationAccess  # service


class HierarchyAccessComp(
    ComponentBase,
    exact_name_partial.ExactNamePartial,
    property_set_info_partial.PropertySetInfoPartial,
    property_state_partial.PropertyStatePartial,
    multi_property_states_partial.MultiPropertyStatesPartial,
    property_with_state_partial.PropertyWithStatePartial,
    child_partial.ChildPartial,
):
    """
    Class for managing ConfigurationAccess Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: Any) -> None:
        """
        Constructor

        Args:
            component (XNameAccess): UNO Component that implements ``com.sun.star.container.XNameAccess``.
        """

        ComponentBase.__init__(self, component)
        exact_name_partial.ExactNamePartial.__init__(self, component=component, interface=None)
        property_set_info_partial.PropertySetInfoPartial.__init__(self, component=component, interface=None)
        property_state_partial.PropertyStatePartial.__init__(self, component=component, interface=None)
        multi_property_states_partial.MultiPropertyStatesPartial.__init__(self, component=component, interface=None)
        property_with_state_partial.PropertyWithStatePartial.__init__(self, component=component, interface=None)
        child_partial.ChildPartial.__init__(self, component=component, interface=None)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.configuration.ConfigurationAccess",)

    # endregion Overrides

    # region Properties

    @property
    def component(self) -> ConfigurationAccess:
        """ConfigurationAccess Component"""
        # pylint: disable=no-member
        return cast("ConfigurationAccess", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties


def get_builder(component: Any, lo_inst: Any = None) -> Any:
    # pylint: disable=import-outside-toplevel
    from ooodev.utils.builder.default_builder import DefaultBuilder

    builder = DefaultBuilder(component, lo_inst)
    builder.add_import(
        name="ooodev.adapter.configuration.hierarchy_access_comp.HierarchyAccessComp",
        uno_name="com.sun.star.configuration.ConfigurationAccess",
        optional=False,
        init_kind=1,
    )
    # region Include builders
    inc_na = cast(DefaultBuilder, name_access_partial.get_builder(component, lo_inst))
    builder.add_from_instance(inc_na, make_optional=True)
    inc_ha = cast(DefaultBuilder, hierarchical_name_access_partial.get_builder(component, lo_inst))
    builder.add_from_instance(inc_ha, make_optional=True)
    inc_cp = cast(DefaultBuilder, container_partial.get_builder(component, lo_inst))
    builder.add_from_instance(inc_cp, make_optional=True)
    inc_hnp = cast(DefaultBuilder, hierarchical_name_partial.get_builder(component, lo_inst))
    builder.add_from_instance(inc_hnp, make_optional=True)
    inc_np = cast(DefaultBuilder, named_partial.get_builder(component, lo_inst))
    builder.add_from_instance(inc_np, make_optional=True)
    inc_pp = cast(DefaultBuilder, property_partial.get_builder(component, lo_inst))
    builder.add_from_instance(inc_pp, make_optional=True)
    # endregion Include builders

    # region exclude builders
    ex_nap = cast(DefaultBuilder, exact_name_partial.get_builder(component, lo_inst))
    builder.set_omit(*ex_nap.get_import_names())
    ex_psp = cast(DefaultBuilder, property_set_info_partial.get_builder(component, lo_inst))
    builder.set_omit(*ex_psp.get_import_names())
    ex_pst = cast(DefaultBuilder, property_state_partial.get_builder(component, lo_inst))
    builder.set_omit(*ex_pst.get_import_names())
    ex_mps = cast(DefaultBuilder, multi_property_states_partial.get_builder(component, lo_inst))
    builder.set_omit(*ex_mps.get_import_names())
    ex_pws = cast(DefaultBuilder, property_with_state_partial.get_builder(component, lo_inst))
    builder.set_omit(*ex_pws.get_import_names())
    ex_cp = cast(DefaultBuilder, child_partial.get_builder(component, lo_inst))
    builder.set_omit(*ex_cp.get_import_names())
    # endregion exclude builders

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
