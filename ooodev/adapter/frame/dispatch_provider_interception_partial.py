from __future__ import annotations
from typing import Any, TYPE_CHECKING, Tuple
import uno

from com.sun.star.frame import XDispatchProviderInterception

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.utils.type_var import UnoInterface

if TYPE_CHECKING:
    from com.sun.star.frame import XDispatchProviderInterceptor


class DispatchProviderInterceptionPartial:
    """
    Partial class for XDispatchProviderInterception.
    """

    def __init__(
        self, component: XDispatchProviderInterception, interface: UnoInterface | None = XDispatchProviderInterception
    ) -> None:
        """
        Constructor

        Args:
            component (XDispatchProviderInterception): UNO Component that implements ``com.sun.star.frame.XDispatchProviderInterception`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XDispatchProviderInterception``.
        """
        self.__interface = interface
        self.__validate(component)
        self.__component = component

    def __validate(self, component: Any) -> None:
        """
        Validates the component.

        Args:
            component (Any): The component to be validated.
        """
        if self.__interface is None:
            return
        if not mLo.Lo.is_uno_interfaces(component, self.__interface):
            raise mEx.MissingInterfaceError(self.__interface)

    # region XDispatchProviderInterception
    def register_dispatch_provider_interceptor(self, interceptor: XDispatchProviderInterceptor) -> None:
        """
        Registers an ``XDispatchProviderInterceptor``, which will become the first interceptor in the chain of registered interceptors.
        """
        self.__component.registerDispatchProviderInterceptor(interceptor)

    def release_dispatch_provider_interceptor(self, interceptor: XDispatchProviderInterceptor) -> None:
        """
        removes an ``XDispatchProviderInterceptor`` which was previously registered

        The order of removals is arbitrary. It is not necessary to remove the last registered interceptor first.
        """
        self.__component.releaseDispatchProviderInterceptor(interceptor)

    # endregion XDispatchProviderInterception
