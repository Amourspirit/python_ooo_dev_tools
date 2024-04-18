from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.frame import XUIControllerRegistration

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class UIControllerRegistrationPartial:
    """
    Partial class for XUIControllerRegistration.
    """

    def __init__(
        self, component: XUIControllerRegistration, interface: UnoInterface | None = XUIControllerRegistration
    ) -> None:
        """
        Constructor

        Args:
            component (XUIControllerRegistration): UNO Component that implements ``com.sun.star.frame.XUIControllerRegistration`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XUIControllerRegistration``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XUIControllerRegistration
    def deregister_controller(self, cmd_url: str, model_name: str) -> None:
        """
        function to remove a previously defined association between a user interface controller implementation and a command URL and optional module.
        """
        self.__component.deregisterController(cmd_url, model_name)

    def has_controller(self, cmd_url: str, model_name: str) -> bool:
        """
        function to check if an user interface controller is registered for a command URL and optional module.
        """
        return self.__component.hasController(cmd_url, model_name)

    def register_controller(self, cmd_url: str, model_name: str, controller_implementation_name: str) -> None:
        """
        function to create an association between a user interface controller implementation and a command URL and optional module.
        """
        self.__component.registerController(cmd_url, model_name, controller_implementation_name)

    # endregion XUIControllerRegistration
