from __future__ import annotations
from re import M
from typing import Any, cast, TYPE_CHECKING, Tuple
import uno

from com.sun.star.lang import XMultiComponentFactory

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.uno import XInterface
    from com.sun.star.uno import XComponentContext
    from ooodev.utils.type_var import UnoInterface


class MultiComponentFactoryPartial:
    """
    Partial class for XMultiComponentFactory.
    """

    def __init__(
        self, component: XMultiComponentFactory, interface: UnoInterface | None = XMultiComponentFactory
    ) -> None:
        """
        Constructor

        Args:
            component (XMultiComponentFactory ): UNO Component that implements ``com.sun.star.lang.XMultiComponentFactory`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XMultiComponentFactory``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XMultiComponentFactory
    def create_instance_with_arguments_and_context(
        self, service_name: str, ctx: XComponentContext, *args: Any
    ) -> XInterface:
        """
        Creates an instance of a component which supports the services specified by the factory, and initializes the new instance with the given arguments and context.

        Raises:
            com.sun.star.uno.Exception: ``Exception``
        """
        return self.__component.createInstanceWithArgumentsAndContext(service_name, args, ctx)

    def create_instance_with_context(self, service_name: str, ctx: XComponentContext) -> XInterface:
        """
        Creates an instance of a component which supports the services specified by the factory.

        Raises:
            com.sun.star.uno.Exception: ``Exception``
        """
        return self.__component.createInstanceWithContext(service_name, ctx)

    def get_available_service_names(self) -> Tuple[str, ...]:
        """
        Gets the names of all supported services.
        """
        return self.__component.getAvailableServiceNames()

    # endregion XMultiComponentFactory
