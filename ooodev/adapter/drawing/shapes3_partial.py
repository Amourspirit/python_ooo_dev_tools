from __future__ import annotations
from typing import Tuple, TYPE_CHECKING
import uno

from com.sun.star.drawing import XShapes3

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class Shapes3Partial:
    """
    Partial class for XShapes3 interface.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XShapes3, interface: UnoInterface | None = XShapes3) -> None:
        """
        Constructor

        Args:
            component (XShapes3): UNO Component that implements ``com.sun.star.drawing.XShapes3`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XShapes3``.
        """
        self.__component = component

    # region XShapes3
    def sort(self, sort_order: Tuple[int, ...]) -> None:
        """
        Sort shapes according to given sort order, for perf reason just rearrange and don't broadcast.

        Args:
            comparator (Any): Desired order of the shapes
        """
        sequence = uno.Any("[]long", sort_order)  # type: ignore
        uno.invoke(self.__component, "sort", sequence)

    # endregion XShapes3
