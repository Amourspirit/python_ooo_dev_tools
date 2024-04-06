from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno
from com.sun.star.util import XRefreshable
from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface
    from com.sun.star.util import XRefreshListener


class RefreshablePartial:
    """
    Partial Class XRefreshable.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XRefreshable, interface: UnoInterface | None = XRefreshable) -> None:
        """
        Constructor

        Args:
            component (XRefreshable): UNO Component that implements ``com.sun.star.util.XRefreshable`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XRefreshable``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XRefreshable
    def add_refresh_listener(self, listener: XRefreshListener) -> None:
        """
        Adds the specified listener to receive the event \"refreshed.\"
        """
        self.__component.addRefreshListener(listener)

    def refresh(self) -> None:
        """
        refreshes the data of the object from the connected data source.
        """
        self.__component.refresh()

    def remove_refresh_listener(self, listener: XRefreshListener) -> None:
        """
        removes the specified listener.
        """
        self.__component.removeRefreshListener(listener)

    # endregion XRefreshable


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
    builder.add_import(
        name="ooodev.adapter.util.refreshable_partial.RefreshablePartial",
        uno_name="com.sun.star.util.XRefreshable",
        optional=False,
        init_kind=2,
    )
    return builder
