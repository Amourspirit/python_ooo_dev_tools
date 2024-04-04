from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING

from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.container import named_partial
from ooodev.adapter.container import hierarchical_name_partial

if TYPE_CHECKING:
    from com.sun.star.configuration import HierarchyElement  # service


class HierarchyElementComp(
    ComponentBase,
    named_partial.NamedPartial,
    hierarchical_name_partial.HierarchicalNamePartial,
):
    """
    Class for managing HierarchyElement Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: Any) -> None:
        """
        Constructor

        Args:
            component (Any): UNO Component that supports ``com.sun.star.configuration.HierarchyElement`` service.
        """

        ComponentBase.__init__(self, component)
        named_partial.NamedPartial.__init__(self, component=component, interface=None)
        hierarchical_name_partial.HierarchicalNamePartial.__init__(self, component=component, interface=None)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.configuration.HierarchyElement",)

    # endregion Overrides

    # region Properties

    @property
    def component(self) -> HierarchyElement:
        """HierarchyElement Component"""
        # pylint: disable=no-member
        return cast("HierarchyElement", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties


def get_builder(component: Any, lo_inst: Any = None, **kwargs) -> Any:
    # pylint: disable=import-outside-toplevel
    from ooodev.utils.builder.default_builder import DefaultBuilder

    builder = DefaultBuilder(component, lo_inst)

    local = kwargs.get("local", False)

    na_np = cast(DefaultBuilder, named_partial.get_builder(component, lo_inst))
    ha_hp = cast(DefaultBuilder, hierarchical_name_partial.get_builder(component, lo_inst))
    builder.omits.update(na_np.omits)
    builder.omits.update(ha_hp.omits)

    if local:
        builder.set_omit(*na_np.get_import_names())
        builder.set_omit(*ha_hp.get_import_names())
    else:
        builder.add_from_instance(na_np, make_optional=True)
        builder.add_from_instance(ha_hp, make_optional=True)

    # in a from_lo method in this class the HierarchyAccessComp would be removed and used as the base class

    builder.auto_add_interface("com.sun.star.beans.XProperty", optional=True)
    builder.auto_add_interface("com.sun.star.beans.XPropertyWithState", optional=True)
    builder.auto_add_interface("com.sun.star.container.XChild", optional=True)
    return builder
