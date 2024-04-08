from __future__ import annotations
from typing import Any, TYPE_CHECKING

import uno
from com.sun.star.frame import XControllerBorder

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo


if TYPE_CHECKING:
    from com.sun.star.frame import XBorderResizeListener
    from com.sun.star.frame import BorderWidths
    from com.sun.star.awt import Rectangle
    from ooodev.utils.type_var import UnoInterface


class ControllerBorderPartial:
    """
    Partial class for XControllerBorder.
    """

    def __init__(self, component: XControllerBorder, interface: UnoInterface | None = XControllerBorder) -> None:
        """
        Constructor

        Args:
            component (XControllerBorder): UNO Component that implements ``com.sun.star.frame.XControllerBorder`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XControllerBorder``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XControllerBorder
    def add_border_resize_listener(self, listener: XBorderResizeListener) -> None:
        """
        Adds the specified listener to receive events about controller's border resizing.
        """
        self.__component.addBorderResizeListener(listener)

    def get_border(self) -> BorderWidths:
        """
        allows to get current border sizes of the document.
        """
        return self.__component.getBorder()

    def query_bordered_area(self, preliminary_rectangle: Rectangle) -> Rectangle:
        """
        Allows to get suggestion for resizing of object area surrounded by the border.

        If the view is going to be resized/moved this method can be used to get suggested object area. Pixels are used as units.
        """
        return self.__component.queryBorderedArea(preliminary_rectangle)

    def removeBorderResizeListener(self, xListener: XBorderResizeListener) -> None:
        """
        Removes the specified listener.
        """
        self.__component.removeBorderResizeListener(xListener)

    # endregion XControllerBorder
