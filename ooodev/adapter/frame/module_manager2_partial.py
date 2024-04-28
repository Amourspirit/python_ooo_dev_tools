from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.frame import XModuleManager2
from ooodev.utils.builder.default_builder import DefaultBuilder
from ooodev.adapter.frame.module_manager_partial import ModuleManagerPartial
from ooodev.adapter.container import name_replace_partial

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class ModuleManager2Partial(ModuleManagerPartial, name_replace_partial.NameReplacePartial[Any]):
    """
    Partial class for XModuleManager2.
    """

    def __init__(self, component: XModuleManager2, interface: UnoInterface | None = XModuleManager2) -> None:
        """
        Constructor

        Args:
            component (XModuleManager2): UNO Component that implements ``com.sun.star.frame.XModuleManager2`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XModuleManager2``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        ModuleManagerPartial.__init__(self, component=component, interface=None)
        name_replace_partial.NameReplacePartial.__init__(self, component=component, interface=None)

        self.__component = component


def get_builder(component: Any) -> DefaultBuilder:
    """
    Get the builder for the component.

    Args:
        component (Any): The component.

    Returns:
        DefaultBuilder: Builder instance.
    """

    builder = DefaultBuilder(component)
    builder.auto_add_interface("com.sun.star.frame.XModuleManager2")
    builder.set_omit("com.sun.star.frame.XModuleManager")
    builder.merge(name_replace_partial.get_builder(component))
    return builder
