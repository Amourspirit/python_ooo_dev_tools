from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.configuration import XTemplateInstance

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class TemplateInstancePartial:
    """
    Partial class for XTemplateInstance.
    """

    def __init__(self, component: XTemplateInstance, interface: UnoInterface | None = XTemplateInstance) -> None:
        """
        Constructor

        Args:
            component (XTemplateInstance): UNO Component that implements ``com.sun.star.configuration.XTemplateInstance`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XTemplateInstance``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XTemplateInstance

    def get_template_name(self) -> str:
        """
        Gets the name of the template
        """
        return self.__component.getTemplateName()

    # endregion XTemplateInstance


def get_builder(component: Any) -> Any:
    """
    Get the builder for the component.

    Args:
        component (Any): The component.

    Returns:
        DefaultBuilder: Builder instance.
    """
    # pylint: disable=import-outside-toplevel
    from ooodev.utils.builder.default_builder import DefaultBuilder

    builder = DefaultBuilder(component)
    builder.auto_add_interface("com.sun.star.configuration.XTemplateInstance", False)
    return builder
