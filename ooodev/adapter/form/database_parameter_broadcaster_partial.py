from __future__ import annotations
from typing import Any, TYPE_CHECKING

from com.sun.star.form import XDatabaseParameterBroadcaster

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.form import XDatabaseParameterListener
    from ooodev.utils.type_var import UnoInterface


class DatabaseParameterBroadcasterPartial:
    """
    Partial Class for ``XDatabaseParameterBroadcaster``.
    """

    def __init__(
        self, component: XDatabaseParameterBroadcaster, interface: UnoInterface | None = XDatabaseParameterBroadcaster
    ) -> None:
        """
        Constructor

        Args:
            component (XDatabaseParameterBroadcaster): UNO Component that implements ``com.sun.star.container.XDatabaseParameterBroadcaster``.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``None``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XDatabaseParameterBroadcaster
    def add_parameter_listener(self, listener: XDatabaseParameterListener) -> None:
        """
        Adds the specified listener, to allow it to fill in necessary parameter values.
        """
        self.__component.addParameterListener(listener)

    def remove_parameter_listener(self, listener: XDatabaseParameterListener) -> None:
        """
        Removes the specified listener.
        """
        self.__component.removeParameterListener(listener)

    # endregion XDatabaseParameterBroadcaster
