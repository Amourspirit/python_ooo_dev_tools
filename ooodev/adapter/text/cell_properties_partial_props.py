from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno

from ooo.dyn.style.graphic_location import GraphicLocation
from ooo.dyn.text.vert_orientation import VertOrientationEnum
from ooodev.adapter.text.text_section_comp import TextSectionComp
from ooodev.adapter.table.border_line2_struct_comp import BorderLine2StructComp
from ooodev.adapter.text.text_comp import TextComp
from ooodev.events.events import Events
from ooodev.utils import info as mInfo
from ooodev.units.unit_mm100 import UnitMM100

if TYPE_CHECKING:
    from com.sun.star.graphic import XGraphic
    from com.sun.star.table import BorderLine
    from com.sun.star.text import CellProperties  # service
    from ooodev.events.args.key_val_args import KeyValArgs
    from ooodev.utils.color import Color
    from ooodev.units.unit_obj import UnitT


class CellPropertiesPartialProps:
    """
    Partial class for CellProperties service.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: CellProperties) -> None:
        """
        Constructor

        Args:
            component (CellProperties): UNO Component that implements ``com.sun.star.text.CellProperties`` service.
        """

        self.__component = component
        self.__event_provider = Events(self)
        self.__props = {}

        def on_comp_struct_changed(src: Any, event_args: KeyValArgs) -> None:
            prop_name = str(event_args.event_data["prop_name"])
            if hasattr(self.__component, prop_name):
                setattr(self.__component, prop_name, event_args.source.component)

        self.__fn_on_comp_struct_changed = on_comp_struct_changed
        # pylint: disable=no-member
        self.__event_provider.subscribe_event(
            "com_sun_star_table_BorderLine_changed", self.__fn_on_comp_struct_changed
        )

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
    def bottom_border(self) -> BorderLine2StructComp:
        """
        Gets/Sets a description of the bottom border line of each cell.

        Setting value can be done with a ``BorderLine`` or ``BorderLine2StructComp`` object.

        Returns:
            BorderLine2StructComp: Returns Border Line.

        Hint:
            - ``BorderLine`` can be imported from ``ooo.dyn.table.border_line``
        """
        key = "BottomBorder"
        prop = self.__props.get(key, None)
        if prop is None:
            prop = BorderLine2StructComp(self.__component.BottomBorder, key, self.__event_provider)  # type: ignore
            self.__props[key] = prop
        return cast(BorderLine2StructComp, prop)

    @bottom_border.setter
    def bottom_border(self, value: BorderLine | BorderLine2StructComp) -> None:
        key = "BottomBorder"
        if mInfo.Info.is_instance(value, BorderLine2StructComp):
            self.__component.BottomBorder = value.copy()
        else:
            self.__component.BottomBorder = cast("BorderLine", value)
        if key in self.__props:
            del self.__props[key]

    @property
    def bottom_border_distance(self) -> UnitMM100:
        """
        Gets/Sets the distance of the bottom border.

        When setting the value, it can be an integer (in ``1/100mm`` units) or a ``UnitT`` object.

        Returns:
            UnitMM100: Returns the distance of the bottom border.
        """
        return UnitMM100(self.__component.BottomBorderDistance)

    @bottom_border_distance.setter
    def bottom_border_distance(self, value: int | UnitT) -> None:
        val = UnitMM100.from_unit_val(value)
        self.__component.BottomBorderDistance = val.value

    @property
    def cell_name(self) -> str:
        """
        Gets the cell name.
        """
        return self.__component.CellName

    # @cell_name.setter
    # def cell_name(self, value: str) -> None:
    #     self.__component.CellName = value

    @property
    def is_protected(self) -> bool:
        """
        Gets/Sets whether the cell is write protected or not.
        """
        return self.__component.IsProtected

    @is_protected.setter
    def is_protected(self, value: bool) -> None:
        self.__component.IsProtected = value

    @property
    def left_border(self) -> BorderLine2StructComp:
        """
        Gets/Sets a description of the left border line of each cell.

        Setting value can be done with a ``BorderLine`` or ``BorderLine2StructComp`` object.

        Returns:
            BorderLine2StructComp: Returns Border Line.

        Hint:
            - ``BorderLine`` can be imported from ``ooo.dyn.table.border_line``
        """
        key = "LeftBorder"
        prop = self.__props.get(key, None)
        if prop is None:
            prop = BorderLine2StructComp(self.__component.LeftBorder, key, self.__event_provider)  # type: ignore
            self.__props[key] = prop
        return cast(BorderLine2StructComp, prop)

    @left_border.setter
    def left_border(self, value: BorderLine | BorderLine2StructComp) -> None:
        key = "LeftBorder"
        if mInfo.Info.is_instance(value, BorderLine2StructComp):
            self.__component.LeftBorder = value.copy()
        else:
            self.__component.LeftBorder = cast("BorderLine", value)
        if key in self.__props:
            del self.__props[key]

    @property
    def left_border_distance(self) -> UnitMM100:
        """
        Gets/Sets the distance of the left border.

        When setting the value, it can be an integer (in ``1/100mm`` units) or a ``UnitT`` object.

        Returns:
            UnitMM100: Returns the distance of the left border.
        """
        return UnitMM100(self.__component.LeftBorderDistance)

    @left_border_distance.setter
    def left_border_distance(self, value: int | UnitT) -> None:
        val = UnitMM100.from_unit_val(value)
        self.__component.LeftBorderDistance = val.value

    @property
    def number_format(self) -> int:
        """
        Gets/Sets the index of the number format that is used in the cells.

        The proper value can be determined by using the ``com.sun.star.util.NumberFormatter`` interface of the document.
        """
        return self.__component.NumberFormat

    @number_format.setter
    def number_format(self, value: int) -> None:
        self.__component.NumberFormat = value

    @property
    def parent_text(self) -> TextComp | None:
        """
        Gets the parent text of this table cell.

        This might be a header text, body text, parent cell, etc.

        **optional**
        """
        if hasattr(self.__component, "ParentText"):
            return TextComp(self.__component.ParentText)
        return None

    @property
    def right_border(self) -> BorderLine2StructComp:
        """
        Gets/Sets a description of the right border line of each cell.

        Setting value can be done with a ``BorderLine`` or ``BorderLine2StructComp`` object.

        Returns:
            BorderLine2StructComp: Returns Border Line.

        Hint:
            - ``BorderLine`` can be imported from ``ooo.dyn.table.border_line``
        """
        key = "RightBorder"
        prop = self.__props.get(key, None)
        if prop is None:
            prop = BorderLine2StructComp(self.__component.RightBorder, key, self.__event_provider)  # type: ignore
            self.__props[key] = prop
        return cast(BorderLine2StructComp, prop)

    @right_border.setter
    def right_border(self, value: BorderLine | BorderLine2StructComp) -> None:
        key = "RightBorder"
        if mInfo.Info.is_instance(value, BorderLine2StructComp):
            self.__component.RightBorder = value.copy()
        else:
            self.__component.RightBorder = cast("BorderLine", value)
        if key in self.__props:
            del self.__props[key]

    @property
    def right_border_distance(self) -> UnitMM100:
        """
        Gets/Sets the distance of the right border.

        When setting the value, it can be an integer (in ``1/100mm`` units) or a ``UnitT`` object.

        Returns:
            UnitMM100: Returns the distance of the right border.
        """
        return UnitMM100(self.__component.RightBorderDistance)

    @right_border_distance.setter
    def right_border_distance(self, value: int | UnitT) -> None:
        val = UnitMM100.from_unit_val(value)
        self.__component.RightBorderDistance = val.value

    @property
    def text_section(self) -> TextSectionComp | None:
        """
        Gets the text section the text table is contained in if there is any.
        """
        section = self.__component.TextSection
        return None if section is None else TextSectionComp(section)

    @property
    def top_border(self) -> BorderLine2StructComp:
        """
        Gets/Sets a description of the top border line of each cell.

        Setting value can be done with a ``BorderLine`` or ``BorderLine2StructComp`` object.

        Returns:
            BorderLine2StructComp: Returns Border Line.

        Hint:
            - ``BorderLine`` can be imported from ``ooo.dyn.table.border_line``
        """
        key = "TopBorder"
        prop = self.__props.get(key, None)
        if prop is None:
            prop = BorderLine2StructComp(self.__component.TopBorder, key, self.__event_provider)  # type: ignore
            self.__props[key] = prop
        return cast(BorderLine2StructComp, prop)

    @top_border.setter
    def top_border(self, value: BorderLine | BorderLine2StructComp) -> None:
        key = "TopBorder"
        if mInfo.Info.is_instance(value, BorderLine2StructComp):
            self.__component.TopBorder = value.copy()
        else:
            self.__component.TopBorder = cast("BorderLine", value)
        if key in self.__props:
            del self.__props[key]

    @property
    def top_border_distance(self) -> UnitMM100:
        """
        Gets/Sets the distance of the top border.

        When setting the value, it can be an integer (in ``1/100mm`` units) or a ``UnitT`` object.

        Returns:
            UnitMM100: Returns the distance of the bottom border.
        """
        return UnitMM100(self.__component.TopBorderDistance)

    @top_border_distance.setter
    def top_border_distance(self, value: int) -> None:
        val = UnitMM100.from_unit_val(value)
        self.__component.TopBorderDistance = val.value

    @property
    def vert_orient(self) -> VertOrientationEnum:
        """
        Gets/Sets the vertical orientation of the text inside of the table cells in this row.

        When setting the value, it can be an integer or an instance of ``VertOrientationEnum``.

        Returns:
            VertOrientationEnum: Returns Vertical Orientation.

        Hint:
            - ``VertOrientationEnum`` can be imported from ``ooo.dyn.text.vert_orientation``
        """
        return VertOrientationEnum(self.__component.VertOrient)

    @vert_orient.setter
    def vert_orient(self, value: int | VertOrientationEnum) -> None:
        val = VertOrientationEnum(value)
        self.__component.VertOrient = val.value

    # endregion Properties
