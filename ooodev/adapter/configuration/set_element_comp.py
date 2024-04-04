from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING

from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.beans import property_with_state_partial
from ooodev.adapter.container import child_partial
from ooodev.adapter.lang import component_partial
from ooodev.adapter.configuration import template_instance_partial

if TYPE_CHECKING:
    from com.sun.star.configuration import SetElement  # service


class SetElementComp(
    ComponentBase,
    property_with_state_partial.PropertyWithStatePartial,
    child_partial.ChildPartial,
    component_partial.ComponentPartial,
    template_instance_partial.TemplateInstancePartial,
):
    """
    Class for managing SetElement Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: Any) -> None:
        """
        Constructor

        Args:
            component (Any): UNO Component that supports ``com.sun.star.configuration.SetElement`` service.
        """

        ComponentBase.__init__(self, component)
        property_with_state_partial.PropertyWithStatePartial.__init__(self, component=component, interface=None)
        child_partial.ChildPartial.__init__(self, component=component, interface=None)
        component_partial.ComponentPartial.__init__(self, component=component, interface=None)
        template_instance_partial.TemplateInstancePartial.__init__(self, component=component, interface=None)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.configuration.SetElement",)

    # endregion Overrides

    # region Properties

    @property
    def component(self) -> SetElement:
        """SetElement Component"""
        # pylint: disable=no-member
        return cast("SetElement", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties


def get_builder(component: Any, lo_inst: Any = None, **kwargs) -> Any:
    # pylint: disable=import-outside-toplevel
    from ooodev.utils.builder.default_builder import DefaultBuilder

    builder = DefaultBuilder(component, lo_inst)

    local = kwargs.get("local", False)

    pwsp = cast(DefaultBuilder, property_with_state_partial.get_builder(component, lo_inst))
    cpp = cast(DefaultBuilder, child_partial.get_builder(component, lo_inst))
    cp = cast(DefaultBuilder, component_partial.get_builder(component, lo_inst))
    tmpl = cast(DefaultBuilder, template_instance_partial.get_builder(component, lo_inst))
    builder.omits.update(pwsp.omits)
    builder.omits.update(cpp.omits)
    builder.omits.update(cp.omits)
    builder.omits.update(tmpl.omits)

    if local:
        builder.set_omit(*pwsp.get_import_names())
        builder.set_omit(*cpp.get_import_names())
        builder.set_omit(*cp.get_import_names())
        builder.set_omit(*tmpl.get_import_names())
    else:
        builder.add_from_instance(pwsp, make_optional=True)
        builder.add_from_instance(cpp, make_optional=True)
        builder.add_from_instance(cp, make_optional=True)
        builder.add_from_instance(tmpl, make_optional=True)

    # in a from_lo method in this class the HierarchyAccessComp would be removed and used as the base class
    builder.auto_add_interface("com.sun.star.container.XHierarchicalName", optional=True)
    builder.auto_add_interface("com.sun.star.container.XNamed", optional=True)
    builder.auto_add_interface("com.sun.star.beans.XProperty", optional=True)

    return builder
