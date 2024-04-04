from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING

from ooodev.adapter.container import child_partial
from ooodev.adapter.configuration import hierarchy_element_comp

if TYPE_CHECKING:
    from com.sun.star.configuration import GroupElement  # service


class GroupElementComp(
    hierarchy_element_comp.HierarchyElementComp,
    child_partial.ChildPartial,
):
    """
    Class for managing GroupElement Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: Any) -> None:
        """
        Constructor

        Args:
            component (Any): UNO Component that supports ``com.sun.star.configuration.GroupElement`` service.
        """

        hierarchy_element_comp.HierarchyElementComp.__init__(self, component)
        child_partial.ChildPartial.__init__(self, component=component, interface=None)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.configuration.GroupElement",)

    # endregion Overrides

    # region Properties

    @property
    def component(self) -> GroupElement:
        """GroupElement Component"""
        # pylint: disable=no-member
        return cast("GroupElement", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties


def get_builder(component: Any, lo_inst: Any = None, **kwargs) -> Any:
    # pylint: disable=import-outside-toplevel
    from ooodev.utils.builder.default_builder import DefaultBuilder

    builder = DefaultBuilder(component, lo_inst)

    local = kwargs.get("local", False)

    hec = cast(DefaultBuilder, hierarchy_element_comp.get_builder(component, lo_inst))
    cpp = cast(DefaultBuilder, child_partial.get_builder(component, lo_inst))
    builder.omits.update(hec.omits)
    builder.omits.update(cpp.omits)

    if local:
        builder.set_omit(*hec.get_import_names())
        builder.set_omit(*cpp.get_import_names())
    else:
        builder.add_from_instance(hec, make_optional=True)
        builder.add_from_instance(cpp, make_optional=True)

    # XChild is already include with HierarchyElementComp but it is optional.
    builder.auto_add_interface("com.sun.star.container.XChild", optional=False)

    return builder
