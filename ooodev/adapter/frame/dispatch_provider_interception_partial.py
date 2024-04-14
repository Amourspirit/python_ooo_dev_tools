from __future__ import annotations
from typing import Any, TYPE_CHECKING, Tuple
import uno

from com.sun.star.frame import XDispatchProviderInterception


from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.utils.type_var import UnoInterface

if TYPE_CHECKING:
    from ooodev.utils.builder.default_builder import DefaultBuilder
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

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XDispatchProviderInterception
    def register_dispatch_provider_interceptor(self, interceptor: XDispatchProviderInterceptor) -> None:
        """
        Registers an ``XDispatchProviderInterceptor``, which will become the first interceptor in the chain of registered interceptors.
        """
        self.__component.registerDispatchProviderInterceptor(interceptor)

    def release_dispatch_provider_interceptor(self, interceptor: XDispatchProviderInterceptor) -> None:
        """
        Removes an ``XDispatchProviderInterceptor`` which was previously registered

        The order of removals is arbitrary. It is not necessary to remove the last registered interceptor first.
        """
        self.__component.releaseDispatchProviderInterceptor(interceptor)

    # endregion XDispatchProviderInterception


def get_builder(component: Any) -> DefaultBuilder:
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
    builder.auto_add_interface("com.sun.star.frame.XDispatchProviderInterception", False)
    return builder
