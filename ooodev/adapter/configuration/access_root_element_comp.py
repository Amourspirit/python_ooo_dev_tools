from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING

from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.beans import property_with_state_partial
from ooodev.adapter.container import child_partial
from ooodev.adapter.lang import component_partial
from ooodev.adapter.util import changes_notifier_partial
from ooodev.adapter.util import changes_events

if TYPE_CHECKING:
    from com.sun.star.configuration import AccessRootElement  # service


class AccessRootElementComp(
    ComponentBase,
    property_with_state_partial.PropertyWithStatePartial,
    child_partial.ChildPartial,
    component_partial.ComponentPartial,
    changes_notifier_partial.ChangesNotifierPartial,
    changes_events.ChangesEvents,
):
    """
    Class for managing AccessRootElement Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: Any) -> None:
        """
        Constructor

        Args:
            component (Any): UNO Component that supports ``com.sun.star.configuration.AccessRootElement`` service.
        """

        ComponentBase.__init__(self, component)
        property_with_state_partial.PropertyWithStatePartial.__init__(self, component=component, interface=None)
        child_partial.ChildPartial.__init__(self, component=component, interface=None)
        component_partial.ComponentPartial.__init__(self, component=component, interface=None)
        changes_notifier_partial.ChangesNotifierPartial.__init__(self, component=component, interface=None)
        changes_events.ChangesEvents.__init__(self, cb=changes_events.on_lazy_cb)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.configuration.AccessRootElement",)

    # endregion Overrides

    # region Properties

    @property
    def component(self) -> AccessRootElement:
        """AccessRootElement Component"""
        # pylint: disable=no-member
        return cast("AccessRootElement", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties


def get_builder(component: Any, lo_inst: Any = None, **kwargs) -> Any:
    # pylint: disable=import-outside-toplevel
    from ooodev.utils.builder.default_builder import DefaultBuilder

    builder = DefaultBuilder(component, lo_inst)

    local = kwargs.get("local", False)

    pwsp = cast(DefaultBuilder, property_with_state_partial.get_builder(component, lo_inst))
    cpp = cast(DefaultBuilder, child_partial.get_builder(component, lo_inst))
    cp = cast(DefaultBuilder, component_partial.get_builder(component, lo_inst))
    cn = cast(DefaultBuilder, changes_notifier_partial.get_builder(component, lo_inst))
    builder.omits.update(pwsp.omits)
    builder.omits.update(cpp.omits)
    builder.omits.update(cp.omits)
    builder.omits.update(cn.omits)

    if local:
        builder.set_omit(*pwsp.get_import_names())
        builder.set_omit(*cpp.get_import_names())
        builder.set_omit(*cp.get_import_names())
        builder.set_omit(*cn.get_import_names())
    else:
        builder.add_from_instance(pwsp, make_optional=True)
        builder.add_from_instance(cpp, make_optional=True)
        builder.add_from_instance(cp, make_optional=True)
        builder.add_from_instance(cn, make_optional=True)

    builder.set_omit("ooodev.adapter.util.changes_events.ChangesEvents")

    # in a from_lo method in this class the HierarchyAccessComp would be removed and used as the base class
    builder.auto_add_interface("com.sun.star.container.XHierarchicalName", optional=True)
    builder.auto_add_interface("com.sun.star.container.XNamed", optional=True)
    builder.auto_add_interface("com.sun.star.beans.XProperty", optional=True)
    builder.auto_add_interface("com.sun.star.lang.XLocalizable", optional=True)

    return builder
