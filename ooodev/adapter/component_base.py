from __future__ import annotations
from typing import Any
from abc import ABC

import uno  # pylint: disable=unused-import
from ooodev.events.args.generic_args import GenericArgs


class ComponentBase(ABC):
    """
    Base Class for Components in the ``component`` name space.
    """

    def __init__(self, component: Any) -> None:
        self.__set_component(component)
        self.__generic_args = None

    def __set_component(self, component: Any) -> None:
        """
        Sets the component.

        Args:
            component (Any): UNO Object

        Raises:
            NotSupportedServiceError: If the component does not support the required service.
        """
        if not self.__get_is_supported(component):
            services = self.__get_supported_service_names()
            if services:
                raise mEx.NotSupportedServiceError(*services)
            else:
                raise mEx.NotSupportedServiceError("No service name specified.")
        self.__component = component

    def __get_component(self) -> Any:
        return self.__component

    def __get_generic_args(self) -> GenericArgs:
        if self.__generic_args is None:
            self.__generic_args = GenericArgs(control_src=self)
        return self.__generic_args

    def __get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ()

    def __get_is_supported(self, component: Any) -> bool:
        """
        Gets whether the component supports a service.

        Args:
            component (component): UNO Object

        Returns:
            bool: True if the component supports the service, otherwise False.
        """
        if component is None:
            return False
        srv_name = self.__get_supported_service_names()
        if not srv_name:
            return True
        return mInfo.Info.support_service(component, *srv_name)


# Leave this import here to avoid circular imports.
from ooodev.utils import info as mInfo
from ooodev.exceptions import ex as mEx
