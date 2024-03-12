from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno
from com.sun.star.drawing import XDrawView

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.drawing import XDrawPage
    from ooodev.utils.type_var import UnoInterface


class DrawViewPartial:
    """Partial class for XDrawView  interface."""

    # Does no implement any methods.
    def __init__(self, component: XDrawView, interface: UnoInterface | None = XDrawView) -> None:
        """
        Constructor

        Args:
            component (XDrawView): UNO Component that implements ``com.sun.star.drawing.XDrawView`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XDrawView``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XDrawView
    def get_current_page(self) -> XDrawPage:
        """
        Gets the current page.
        """
        return self.__component.getCurrentPage()

    def set_current_page(self, page: XDrawPage) -> None:
        """
        Sets the current page.
        """
        self.__component.setCurrentPage(page)

    # endregion XDrawView
