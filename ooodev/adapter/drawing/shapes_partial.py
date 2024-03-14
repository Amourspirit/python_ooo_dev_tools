from __future__ import annotations
from typing import TYPE_CHECKING
import uno

from com.sun.star.drawing import XShapes

from ooodev.adapter.container.index_access_partial import IndexAccessPartial

if TYPE_CHECKING:
    from com.sun.star.drawing import XShape
    from ooodev.utils.type_var import UnoInterface


class ShapesPartial(IndexAccessPartial["XShape"]):
    """
    Partial class for XShape interface.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XShapes, interface: UnoInterface | None = XShapes) -> None:
        """
        Constructor

        Args:
            component (XShapes): UNO Component that implements ``com.sun.star.drawing.XShapes`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XShapes``.
        """
        self.__component = component
        IndexAccessPartial.__init__(self, component, interface)

    # region XShape
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

    # endregion XShape
