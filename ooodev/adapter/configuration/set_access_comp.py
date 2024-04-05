from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno
from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.container import element_access_partial
from ooodev.adapter.beans import exact_name_partial
from ooodev.adapter.beans import property_set_info_partial
from ooodev.adapter.beans import property_state_partial
from ooodev.adapter.beans import multi_property_states_partial
from ooodev.adapter.container import container_partial


if TYPE_CHECKING:
    from com.sun.star.configuration import SetAccess  # service


class SetAccessComp(
    ComponentBase,
    exact_name_partial.ExactNamePartial,
    property_set_info_partial.PropertySetInfoPartial,
    property_state_partial.PropertyStatePartial,
    multi_property_states_partial.MultiPropertyStatesPartial,
    container_partial.ContainerPartial,
):
    """
    Class for managing SetAccess Component.
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
        container_partial.ContainerPartial.__init__(self, component=component, interface=None)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.configuration.SetAccess",)

    # endregion Overrides

    # region Properties

    @property
    def component(self) -> SetAccess:
        """SetAccess Component"""
        # pylint: disable=no-member
        return cast("SetAccess", self._ComponentBase__get_component())  # type: ignore

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
    inc_cp = cast(DefaultBuilder, container_partial.get_builder(component, lo_inst))
    builder.omits.update(inc_exact.omits)
    builder.omits.update(inc_psi.omits)
    builder.omits.update(inc_pss.omits)
    builder.omits.update(inc_mps.omits)
    builder.omits.update(inc_cp.omits)

    if local:
        builder.set_omit(*inc_exact.get_import_names())
        builder.set_omit(*inc_psi.get_import_names())
        builder.set_omit(*inc_pss.get_import_names())
        builder.set_omit(*inc_mps.get_import_names())
        builder.set_omit(*inc_cp.get_import_names())
    else:
        builder.add_from_instance(inc_exact, make_optional=True)
        builder.add_from_instance(inc_psi, make_optional=True)
        builder.add_from_instance(inc_pss, make_optional=True)
        builder.add_from_instance(inc_mps, make_optional=True)
        builder.add_from_instance(inc_cp, make_optional=True)

    # endregion exclude local builders

    # region exclude other builders

    ex_el = cast(DefaultBuilder, element_access_partial.get_builder(component, lo_inst))
    builder.set_omit(*ex_el.get_import_names())
    # endregion exclude other builders

    builder.auto_add_interface("com.sun.star.container.XNameAccess")
    builder.auto_add_interface("com.sun.star.container.XHierarchicalNameAccess")
    builder.auto_add_interface("com.sun.star.configuration.XTemplateContainer")
    builder.auto_add_interface("com.sun.star.util.XStringEscape")
    return builder
