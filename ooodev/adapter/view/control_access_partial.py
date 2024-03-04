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

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

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
