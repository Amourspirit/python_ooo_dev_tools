from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno
from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.container import name_access_partial
from ooodev.adapter.container import element_access_partial


if TYPE_CHECKING:
    from com.sun.star.configuration import SimpleSetAccess  # service


class SimpleSetAccessComp(ComponentBase, name_access_partial.NameAccessPartial):
    """
    Class for managing SimpleSetAccess Component.
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

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.configuration.SimpleSetAccess",)

    # endregion Overrides

    # region Properties

    @property
    def component(self) -> SimpleSetAccess:
        """SimpleSetAccess Component"""
        # pylint: disable=no-member
        return cast("SimpleSetAccess", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties


def get_builder(component: Any, lo_inst: Any = None, **kwargs) -> Any:
    # pylint: disable=import-outside-toplevel
    from ooodev.utils.builder.default_builder import DefaultBuilder

    builder = DefaultBuilder(component, lo_inst)
    # when local this modules class is added as the base class.
    # When not local this modules base class in not included but all of its import classer are.
    # see from_lo() above.
    local = kwargs.get("local", False)

    # region exclude local builders
    inc_nap = cast(DefaultBuilder, name_access_partial.get_builder(component, lo_inst))
    builder.omits.update(inc_nap.omits)

    if local:
        builder.set_omit(*inc_nap.get_import_names())
    else:
        builder.add_from_instance(inc_nap, make_optional=True)

    # endregion exclude local builders

    # region exclude other builders

    ex_el = cast(DefaultBuilder, element_access_partial.get_builder(component, lo_inst))
    builder.set_omit(*ex_el.get_import_names())
    # endregion exclude other builders

    builder.add_import(
        name="ooodev.adapter.configuration.template_container_partial.TemplateContainerPartial",
        uno_name="com.sun.star.configuration.XTemplateContainer",
        optional=True,
    )
    builder.add_import(
        name="ooodev.adapter.util.string_escape_partial.StringEscapePartial",
        uno_name="com.sun.star.util.XStringEscape",
        optional=True,
    )
    builder.add_import(
        name="ooodev.adapter.container.container_partial.ContainerPartial",
        uno_name="com.sun.star.container.XContainer",
        optional=True,
    )
    return builder
