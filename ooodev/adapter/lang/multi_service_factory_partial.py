from __future__ import annotations
from typing import Any, TYPE_CHECKING, Tuple
import uno

from com.sun.star.lang import XMultiServiceFactory

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.uno import XInterface
    from ooodev.utils.type_var import UnoInterface


class MultiServiceFactoryPartial:
    """
    Partial class for ``XMultiServiceFactory``.
    """

    def __init__(self, component: XMultiServiceFactory, interface: UnoInterface | None = XMultiServiceFactory) -> None:
        """
        Constructor

        Args:
            component (XMultiServiceFactory  ): UNO Component that implements ``com.sun.star.lang.XMultiServiceFactory  `` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XMultiServiceFactory  ``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XMultiServiceFactory
    def create_instance(self, service_name: str) -> XInterface:
        """
        Creates an instance classified by the specified name.

        Raises:
            com.sun.star.uno.Exception: ``Exception``
        """
        return self.__component.createInstance(service_name)

    def create_instance_with_arguments(self, service_name: str, *args: Any) -> XInterface:
        """
        Creates an instance classified by the specified name and passes the arguments to that instance.

        Raises:
            com.sun.star.uno.Exception: ``Exception``
        """
        return self.__component.createInstanceWithArguments(service_name, args)

    def get_available_service_names(self) -> Tuple[str, ...]:
        """
        Provides the available names of the factory to be used to create instances.
        """
        return self.__component.getAvailableServiceNames()

    # endregion XMultiServiceFactory
