from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno
from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.beans import exact_name_partial
from ooodev.adapter.beans import property_set_info_partial
from ooodev.adapter.beans import property_state_partial
from ooodev.adapter.beans import multi_property_states_partial
from ooodev.adapter.beans import property_set_partial
from ooodev.adapter.beans import multi_property_set_partial
from ooodev.adapter.beans import hierarchical_property_set_partial
from ooodev.adapter.beans import multi_hierarchical_property_set_partial


if TYPE_CHECKING:
    from com.sun.star.configuration import GroupAccess  # service


class GroupAccessComp(
    ComponentBase,
    exact_name_partial.ExactNamePartial,
    property_set_info_partial.PropertySetInfoPartial,
    property_state_partial.PropertyStatePartial,
    multi_property_states_partial.MultiPropertyStatesPartial,
    property_set_partial.PropertySetPartial,
    multi_property_set_partial.MultiPropertySetPartial,
    hierarchical_property_set_partial.HierarchicalPropertySetPartial,
    multi_hierarchical_property_set_partial.MultiHierarchicalPropertySetPartial,
):
    """
    Class for managing GroupAccess Component.
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
        property_set_partial.PropertySetPartial.__init__(self, component=component, interface=None)
        multi_property_set_partial.MultiPropertySetPartial.__init__(self, component=component, interface=None)
        hierarchical_property_set_partial.HierarchicalPropertySetPartial.__init__(
            self, component=component, interface=None
        )
        multi_hierarchical_property_set_partial.MultiHierarchicalPropertySetPartial.__init__(
            self, component=component, interface=None
        )

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.configuration.GroupAccess",)

    # endregion Overrides

    # region Properties

    @property
    def component(self) -> GroupAccess:
        """GroupAccess Component"""
        # pylint: disable=no-member
        return cast("GroupAccess", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties


def get_builder(component: Any, lo_inst: Any = None, **kwargs) -> Any:
    # pylint: disable=import-outside-toplevel
    from ooodev.utils.builder.default_builder import DefaultBuilder

    builder = DefaultBuilder(component, lo_inst)
    # when local this modules class is added as the base class.
    # When not local this modules base class in not included but all of its import classes are.
    # see from_lo() above.
    local = kwargs.get("local", False)

    # region exclude local builders
    inc_exact = cast(DefaultBuilder, exact_name_partial.get_builder(component, lo_inst))
    inc_psi = cast(DefaultBuilder, property_set_info_partial.get_builder(component, lo_inst))
    inc_pss = cast(DefaultBuilder, property_state_partial.get_builder(component, lo_inst))
    inc_mps = cast(DefaultBuilder, multi_property_states_partial.get_builder(component, lo_inst))
    inc_pps = cast(DefaultBuilder, property_set_partial.get_builder(component, lo_inst))
    inc_hps = cast(DefaultBuilder, hierarchical_property_set_partial.get_builder(component, lo_inst))
    inc_mhps = cast(DefaultBuilder, multi_hierarchical_property_set_partial.get_builder(component, lo_inst))

    builder.omits.update(inc_exact.omits)
    builder.omits.update(inc_psi.omits)
    builder.omits.update(inc_pss.omits)
    builder.omits.update(inc_mps.omits)
    builder.omits.update(inc_pps.omits)
    builder.omits.update(inc_hps.omits)
    builder.omits.update(inc_mhps.omits)

    if local:
        builder.set_omit(*inc_exact.get_import_names())
        builder.set_omit(*inc_psi.get_import_names())
        builder.set_omit(*inc_pss.get_import_names())
        builder.set_omit(*inc_mps.get_import_names())
        builder.set_omit(*inc_pps.get_import_names())
        builder.set_omit(*inc_hps.get_import_names())
        builder.set_omit(*inc_mhps.get_import_names())
    else:
        builder.add_from_instance(inc_exact, make_optional=True)
        builder.add_from_instance(inc_psi, make_optional=True)
        builder.add_from_instance(inc_pss, make_optional=True)
        builder.add_from_instance(inc_mps, make_optional=True)
        builder.add_from_instance(inc_pps, make_optional=True)
        builder.add_from_instance(inc_hps, make_optional=True)
        builder.add_from_instance(inc_mhps, make_optional=True)

    # endregion exclude local builders

    # region com.sun.star.configuration.HierarchyAccess optional
    builder.auto_add_interface("com.sun.star.container.XNameAccess", optional=True)
    builder.auto_add_interface("com.sun.star.container.XHierarchicalNameAccess", optional=True)
    builder.auto_add_interface("com.sun.star.container.XContainer", optional=True)
    # endregion com.sun.star.configuration.HierarchyAccess optional

    builder.auto_add_interface("com.sun.star.beans.XMultiPropertyStates", optional=True)

    return builder
