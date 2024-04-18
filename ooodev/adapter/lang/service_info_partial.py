from __future__ import annotations
from typing import Any, TYPE_CHECKING, Tuple
import uno

from com.sun.star.lang import XServiceInfo

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class ServiceInfoPartial:
    """
    Partial class for ``XServiceInfo``.
    """

    def __init__(self, component: XServiceInfo, interface: UnoInterface | None = XServiceInfo) -> None:
        """
        Constructor

        Args:
            component (XServiceInfo): UNO Component that implements ``com.sun.star.lang.XServiceInfo`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XServiceInfo``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XServiceInfo
    def get_implementation_name(self) -> str:
        """
        Provides the implementation name of the service implementation.
        """
        return self.__component.getImplementationName()

    def get_supported_service_names(self) -> Tuple[str, ...]:
        """
        Provides the supported service names of the implementation, including also indirect service names.
        """
        return self.__component.getSupportedServiceNames()

    def supports_service(self, *name: str) -> bool:
        """
        Tests whether any of the specified service(s) are supported.

        Args:
            name (str): One or more service name(s) to test such as ``com.sun.star.awt.MenuBar``.

        Returns:
            bool: ``True`` if any of the specified service(s) are supported; Otherwise, ``False``.
        """
        if not name:
            return False
        for n in name:
            if self.__component.supportsService(n):
                return True

        return False

    def supports_all_services(self, *name: str) -> bool:
        """
        Tests whether all the specified services are supported, i.e.

        Args:
            name (str): One or more service name(s) to test such as ``com.sun.star.awt.MenuBar``.

        Raises:
            ValueError: If no service name is provided.

        Returns:
            bool: ``True`` if all the specified services are supported; Otherwise, ``False``.
        """
        if not name:
            raise ValueError("At least one service name must be provided.")
        for n in name:
            if not self.__component.supportsService(n):
                return False

        return True

    # endregion XServiceInfo


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
    builder.auto_add_interface("com.sun.star.lang.XServiceInfo", False)
    return builder
