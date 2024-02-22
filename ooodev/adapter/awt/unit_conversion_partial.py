from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.awt import XUnitConversion
from ooo.dyn.awt.point import Point
from ooo.dyn.awt.size import Size
from ooo.dyn.util.measure_unit import MeasureUnitEnum

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class UnitConversionPartial:
    """
    Partial class for XUnitConversion.
    """

    def __init__(self, component: XUnitConversion, interface: UnoInterface | None = XUnitConversion) -> None:
        """
        Constructor

        Args:
            component (XUnitConversion): UNO Component that implements ``com.sun.star.awt.XUnitConversion`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XUnitConversion``.
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

    # region XUnitConversion
    def convert_point_to_logic(self, point: Point, target_unit: int | MeasureUnitEnum) -> Point:
        """
        Converts the given Point, which is specified in pixels, into the given logical unit.

        Raises:
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``.

        Returns:
            Point: Converted Point

        Hint:
            - ``MeasureUnitEnum`` can be imported from ``ooo.dyn.util.measure_unit``
        """
        return self.__component.convertPointToLogic(point, int(target_unit))

    def convert_point_to_pixel(self, point: Point, source_unit: int | MeasureUnitEnum) -> Point:
        """
        Converts the given Point, which is specified in the given logical unit, into pixels.

        Raises:
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``.

        Returns:
            Point: Converted Point

        Hint:
            - ``MeasureUnitEnum`` can be imported from ``ooo.dyn.util.measure_unit``
        """
        return self.__component.convertPointToPixel(point, int(source_unit))

    def convert_size_to_logic(self, size: Size, target_unit: int | MeasureUnitEnum) -> Size:
        """
        Converts the given Size, which is specified in pixels, into the given logical unit.

        Raises:
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``.

        Returns:
            Point: Converted Point

        Hint:
            - ``MeasureUnitEnum`` can be imported from ``ooo.dyn.util.measure_unit``
        """
        return self.__component.convertSizeToLogic(size, int(target_unit))

    def convert_size_to_pixel(self, size: Size, source_unit: int | MeasureUnitEnum) -> Size:
        """
        Converts the given Size, which is specified in the given logical unit, into pixels.

        Raises:
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``.

        Returns:
            Point: Converted Point

        Hint:
            - ``MeasureUnitEnum`` can be imported from ``ooo.dyn.util.measure_unit``
        """
        return self.__component.convertSizeToPixel(size, int(source_unit))

    # endregion XUnitConversion
