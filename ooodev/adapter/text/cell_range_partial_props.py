from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno

from ooo.dyn.style.graphic_location import GraphicLocation

if TYPE_CHECKING:
    from com.sun.star.graphic import XGraphic
    from com.sun.star.text import CellRange  # service
    from ooodev.utils.color import Color


class CellRangePartialProps:
    """
    Partial class for CellRange service.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: CellRange) -> None:
        """
        Constructor

        Args:
            component (CellRange): UNO Component that implements ``com.sun.star.text.CellRange`` service.
        """

        self.__component = component

    # region Properties
    @property
    def back_color(self) -> Color:
        """
        Gets/Sets the background color.

        Returns:
            ~ooo.dyn.utils.color.Color: Returns Color.
        """
        return self.__component.BackColor  # type: ignore

    @back_color.setter
    def back_color(self, value: Color) -> None:
        self.__component.BackColor = value  # type: ignore

    @property
    def back_graphic(self) -> XGraphic:
        """
        Gets/Sets the graphic object that is displayed as background graphic.
        """
        return self.__component.BackGraphic

    @back_graphic.setter
    def back_graphic(self, value: XGraphic) -> None:
        self.__component.BackGraphic = value

    @property
    def back_graphic_filter(self) -> str:
        """
        Gets/Sets the name of the graphic filter of the background graphic.
        """
        return self.__component.BackGraphicFilter

    @back_graphic_filter.setter
    def back_graphic_filter(self, value: str) -> None:
        self.__component.BackGraphicFilter = value

    @property
    def back_graphic_location(self) -> GraphicLocation:
        """
        Gets/Sets the position of the background graphic.

        Returns:
            GraphicLocation: Returns GraphicLocation.

        Hint:
            - ``GraphicLocation`` can be imported from ``ooo.dyn.style.graphic_location``
        """
        return self.__component.BackGraphicLocation  # type: ignore

    @back_graphic_location.setter
    def back_graphic_location(self, value: GraphicLocation) -> None:
        self.__component.BackGraphicLocation = value  # type: ignore

    @property
    def back_graphic_url(self) -> str:
        """
        Gets/Sets the URL to the background graphic.

        Returns:
            str: Returns URL to the background graphic.

        Note:
            the new behavior since it this was deprecated: This property can only be set and only external URLs are supported (no more vnd.sun.star.GraphicObject scheme). When an URL is set, then it will load the graphic and set the BackGraphic property.
        """
        result = self.__component.BackGraphicURL
        return "" if result is None else result

    @back_graphic_url.setter
    def back_graphic_url(self, value: str) -> None:
        self.__component.BackGraphicURL = value

    @property
    def back_transparent(self) -> bool:
        """
        Gets/Sets whether the background is transparent.
        """
        return self.__component.BackTransparent

    @back_transparent.setter
    def back_transparent(self, value: bool) -> None:
        self.__component.BackTransparent = value

    @property
    def chart_column_as_label(self) -> bool:
        """
        Gets/Sets if the first column of the table should be treated as axis labels when a chart is to be created.
        """
        return self.__component.ChartColumnAsLabel

    @chart_column_as_label.setter
    def chart_column_as_label(self, value: bool) -> None:
        self.__component.ChartColumnAsLabel = value

    @property
    def chart_row_as_label(self) -> bool:
        """
        Gets/Sets if the first row of the table should be treated as axis labels when a chart is to be created.
        """
        return self.__component.ChartRowAsLabel

    @chart_row_as_label.setter
    def chart_row_as_label(self, value: bool) -> None:
        self.__component.ChartRowAsLabel = value

    @property
    def number_format(self) -> int:
        """
        Gets/Sets the number format.
        """
        return self.__component.NumberFormat

    @number_format.setter
    def number_format(self, value: int) -> None:
        self.__component.NumberFormat = value

    # endregion Properties
