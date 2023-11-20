from __future__ import annotations
from typing import Any
from abc import ABC

import uno  # pylint: disable=unused-import
from com.sun.star.lang import XComponent
from ooodev.events.args.generic_args import GenericArgs
from ooodev.utils import info as mInfo
from ooodev.exceptions import ex as mEx


class ComponentBase(ABC):
    """
    Base Class for Components in the ``component`` name space.
    """

    def __init__(self, component: Any) -> None:
        self._set_component(component)

    def _set_component(self, component: XComponent) -> None:
        """
        Sets the component.

        Args:
            component (XComponent): UNO Object

        Raises:
            TypeError: If the component does not support the required service.
        """
        if not self._get_is_supported(component):
            raise TypeError(f"Component does not support the required service.")
        self._component = component

    def _get_component(self) -> XComponent:
        return self._component

    def _get_generic_args(self) -> GenericArgs:
        try:
            return self.__generic_args
        except AttributeError:
            self.__generic_args = GenericArgs(control_src=self)
            return self.__generic_args

    def _get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ()

    def _get_is_supported(self, component: XComponent) -> bool:
        """
        Gets whether the component supports a service.

        Args:
            component (component): UNO Object

        Returns:
            bool: True if the component supports the service, otherwise False.
        """
        if not component:
            return False
        srv_name = self._get_supported_service_names()
        if not srv_name:
            return True
        return mInfo.Info.support_service(component, *srv_name)

    def __setattr__(self, name: str, value: Any) -> None:
        if name == "_component":
            super().__setattr__(name, value)
        else:
            comp = self._get_component()
            if hasattr(comp, name):
                setattr(comp, name, value)
                return
        raise AttributeError(name)

    def __getattr__(self, name: str) -> Any:
        comp = self._get_component()
        if hasattr(comp, name):
            return getattr(comp, name)
        raise AttributeError(name)
