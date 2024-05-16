from __future__ import annotations
from typing import Any, TYPE_CHECKING

import uno
from com.sun.star.frame import XTitle

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo


if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class TitlePartial:
    """
    Partial class for XTitle.
    """

    def __init__(self, component: XTitle, interface: UnoInterface | None = XTitle) -> None:
        """
        Constructor

        Args:
            component (XTitle): UNO Component that implements ``com.sun.star.frame.XTitle`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XTitle``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XTitle
    def get_title(self) -> str:
        """
        Returns the title of the object.
        """
        return self.__component.getTitle()

    def set_title(self, title: str) -> None:
        """
        Sets the title of the object.
        """
        self.__component.setTitle(title)

    # endregion XTitle
