from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.chart2 import XTitled

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.chart2 import XTitle
    from ooodev.utils.type_var import UnoInterface


class TitledPartial:
    """
    Partial class for XTitled.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XTitled, interface: UnoInterface | None = XTitled) -> None:
        """
        Constructor

        Args:
            component (XTitled): UNO Component that implements ``com.sun.star.chart2.XTitled`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XTitled``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XTitled
    def get_title_object(self) -> XTitle:
        """
        get the object holding the title's content and formatting
        """
        return self.__component.getTitleObject()

    def set_title_object(self, title: XTitle) -> None:
        """
        set a new title object replacing the former one
        """
        self.__component.setTitleObject(title)

    # endregion XTitled
