from __future__ import annotations
from typing import Any, TYPE_CHECKING

import uno
from com.sun.star.view import XControlAccess

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.awt import XControl
    from com.sun.star.awt import XControlModel
    from ooodev.utils.type_var import UnoInterface


class ControlAccessPartial:
    """
    Partial class for XControlAccess.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XControlAccess, interface: UnoInterface | None = XControlAccess) -> None:
        """
        Constructor

        Args:
            component (XControlAccess ): UNO Component that implements ``com.sun.star.view.XControlAccess`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XControlAccess``.
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

    # region XControlAccess
    def get_control(self, model: XControlModel) -> XControl:
        """
        Returns the control from the specified model.

        Args:
            model (XControlModel): The model of the control to be returned.

        Returns:
            XControl: The control from the specified model.
        """
        return self.__component.getControl(model)

    # endregion XControlAccess
