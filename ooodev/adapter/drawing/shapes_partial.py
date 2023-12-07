from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno

if TYPE_CHECKING:
    from com.sun.star.drawing import XShapes
    from com.sun.star.drawing import XShape

from ooodev.utils import lo as mLo
from ooodev.exceptions import ex as mEx
from ooodev.adapter.container.index_access_partial import IndexAccessPartial


class ShapesPartial(IndexAccessPartial):
    """
    Class for managing IndexAccess.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XShapes) -> None:
        """
        Constructor

        Args:
            component (XShapes): UNO Component that implements ``com.sun.star.drawing.XShapes`` interface.
        """
        if not mLo.Lo.is_uno_interfaces(component, XShapes):
            raise mEx.MissingInterfaceError("XShapes")
        IndexAccessPartial.__init__(self, component)
        self.__component = component

    # region Methods
    def add(self, shape: XShape) -> None:
        """
        Adds a shape to the collection.

        Args:
            shape (XShape): The shape to add.
        """
        self.__component.add(shape)

    def remove(self, shape: XShape) -> None:
        """
        Removes a shape from the collection.

        Args:
            shape (XShape): The shape to remove.
        """
        self.__component.remove(shape)

    # endregion Methods
