from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.configuration import XTemplateContainer

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class TemplateContainerPartial:
    """
    Partial class for XTemplateContainer.
    """

    def __init__(self, component: XTemplateContainer, interface: UnoInterface | None = XTemplateContainer) -> None:
        """
        Constructor

        Args:
            component (XTemplateContainer): UNO Component that implements ``com.sun.star.configuration.XTemplateContainer`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XTemplateContainer``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XTemplateContainer

    def get_element_template_name(self) -> str:
        """
        Gets the name of the template

        If instances of multiple templates are accepted by the container, this is the name of the basic or primary template.

        Instances of the template must be created using an appropriate factory.
        """
        return self.__component.getElementTemplateName()

    # endregion XTemplateContainer


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
    builder.auto_add_interface("com.sun.star.configuration.XTemplateContainer", False)
    return builder
