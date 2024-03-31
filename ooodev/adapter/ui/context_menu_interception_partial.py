from __future__ import annotations
from typing import Any, TYPE_CHECKING

from com.sun.star.ui import XContextMenuInterception

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.ui import XContextMenuInterceptor
    from ooodev.utils.type_var import UnoInterface

# https://wiki.documentfoundation.org/Framework/Tutorial/Context_Menu_Interception


class ContextMenuInterceptionPartial:
    """
    Partial Class for XContextMenuInterception.

    .. versionadded:: 0.20.0
    """

    def __init__(
        self, component: XContextMenuInterception, interface: UnoInterface | None = XContextMenuInterception
    ) -> None:
        """
        Constructor

        Args:
            component (XContextMenuInterception): UNO Component that implements ``com.sun.star.ui.XContextMenuInterception``.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XContextMenuInterception``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XContextMenuInterception
    def register_context_menu_interceptor(self, interceptor: XContextMenuInterceptor) -> None:
        """
        Registers a context menu interceptor,
        which will become the first interceptor in the chain of registered interceptors.

        Args:
            interceptor (XContextMenuInterceptor): The interceptor to be registered.
        """
        self.__component.registerContextMenuInterceptor(interceptor)

    def release_context_menu_interceptor(self, interceptor: XContextMenuInterceptor) -> None:
        """
        Releases a context menu interceptor.

        The order of removals is arbitrary. It is not necessary to remove the last registered interceptor first.

        Args:
            interceptor (XContextMenuInterceptor): The interceptor to be released.
        """
        self.__component.releaseContextMenuInterceptor(interceptor)

    # endregion XContextMenuInterception
