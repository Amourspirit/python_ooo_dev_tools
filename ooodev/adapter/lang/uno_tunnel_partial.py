from __future__ import annotations
from typing import Any, TYPE_CHECKING, Tuple
import uno

from com.sun.star.lang import XUnoTunnel

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class UnoTunnelPartial:
    """
    Partial class for XUnoTunnel.

    See Also:
        `API XUnoTunnel <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1lang_1_1XUnoTunnel.html>`_
    """

    def __init__(self, component: XUnoTunnel, interface: UnoInterface | None = XUnoTunnel) -> None:
        """
        Constructor

        Args:
            component (XUnoTunnel): UNO Component that implements ``com.sun.star.lang.XUnoTunnel`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XUnoTunnel``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XUnoTunnel
    def get_something(self, identifier: Tuple[int, ...]) -> int:
        """
        Call this method to get something which is not specified in UNO, e.g.

        an address to some C++ object.
        """
        return self.__component.getSomething(identifier)  # type: ignore

    # endregion XUnoTunnel


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
    builder.auto_add_interface("com.sun.star.lang.XUnoTunnel", False)
    return builder
