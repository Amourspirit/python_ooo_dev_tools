from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING, Tuple
import contextlib
import uno
from ooo.dyn.style.graphic_location import GraphicLocation
from ooo.dyn.style.break_type import BreakType
from ooo.dyn.text.hori_orientation import HoriOrientationEnum
from ooodev.adapter.table.shadow_format_struct_comp import ShadowFormatStructComp
from ooodev.adapter.table.table_border_struct_comp import TableBorderStructComp
from ooodev.events.events import Events
from ooodev.utils import info as mInfo

if TYPE_CHECKING:
    from com.sun.star.text import TextTable  # service
    from com.sun.star.text import TableColumnSeparator  # struct
    from com.sun.star.beans import PropertyValue
    from com.sun.star.graphic import XGraphic
    from com.sun.star.table import ShadowFormat
    from com.sun.star.table import TableBorder
    from ooodev.events.args.key_val_args import KeyValArgs
    from ooodev.utils.color import Color


class TextTablePropertiesPartial:
    """
    Partial Properties class for TextTable Service.

    See Also:
        `API TextTable <https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1text_1_1TextTable.html>`__
    """

    def __init__(self, component: TextTable) -> None:
        """
        Constructor

        Args:
            component (TextTable): UNO Component that implements ``com.sun.star.text.TextTable`` service.
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
            "com_sun_star_table_ShadowFormat_changed", self.__fn_on_comp_struct_changed
        )

        self.__event_provider.subscribe_event(
            "com_sun_star_table_TableBorder_changed", self.__fn_on_comp_struct_changed
        )

    # region Properties

    @property
    def table_column_separators(self) -> Tuple[TableColumnSeparator, ...]:
        """
        Gets/Sets the column description of the table.
        """
        return self.__component.TableColumnSeparators

    @table_column_separators.setter
    def table_column_separators(self, value: Tuple[TableColumnSeparator, ...]) -> None:
        self.__init__.TableColumnSeparators = value

    @property
    def table_interop_grab_bag(self) -> Tuple[PropertyValue, ...] | None:
        """
        Gets/Sets - Grab bag of table properties, used as a string-any map for interim interop purposes.

        This property is intentionally not handled by the ODF filter. Any member that should be handled there should be first moved out from this grab bag to a separate property.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.TableInteropGrabBag
        return None

    @table_interop_grab_bag.setter
    def table_interop_grab_bag(self, value: Tuple[PropertyValue, ...]) -> None:
        with contextlib.suppress(AttributeError):
            self.__init__.TableInteropGrabBag = value

    @property
    def back_color(self) -> Color:
        """
        Gets/Sets the color of the background.

        Returns:
            ~ooodev.utils.color.Color: Color
        """
        return self.__component.BackColor  # type: ignore

    @back_color.setter
    def back_color(self, value: Color) -> None:
        self.__component.BackColor = value  # type: ignore

    @property
    def back_graphic(self) -> XGraphic | None:
        """
        Gets/Sets the graphic for the background.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.BackGraphic
        return None

    @back_graphic.setter
    def back_graphic(self, value: XGraphic) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.BackGraphic = value

    @property
    def back_graphic_filter(self) -> str:
        """
        Gets/Sets the name of the file filter for the background graphic.
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
            GraphicLocation: Graphic Location

        Hint:
            - ``GraphicLocation`` can be imported from ``ooo.dyn.style.graphic_location ``
        """
        return self.__component.BackGraphicLocation  # type: ignore

    @back_graphic_location.setter
    def back_graphic_location(self, value: GraphicLocation) -> None:
        self.__component.BackGraphicLocation = value  # type: ignore

    @property
    def back_graphic_url(self) -> str:
        """
        Gets/Sets the URL for the background graphic.

        Note the new behavior since it this was deprecated: This property can only be set and only external URLs are supported (no more vnd.sun.star.GraphicObject scheme). When an URL is set, then it will load the graphic and set the BackGraphic property.
        """
        return self.__component.BackGraphicURL

    @back_graphic_url.setter
    def back_graphic_url(self, value: str) -> None:
        self.__component.BackGraphicURL = value

    @property
    def back_transparent(self) -> bool:
        """
        Gets/Sets if the background color is transparent.
        """
        return self.__component.BackTransparent

    @back_transparent.setter
    def back_transparent(self, value: bool) -> None:
        self.__component.BackTransparent = value

    @property
    def bottom_margin(self) -> int:
        """
        Gets/Sets the bottom margin.
        """
        return self.__component.BottomMargin

    @bottom_margin.setter
    def bottom_margin(self, value: int) -> None:
        self.__component.BottomMargin = value

    @property
    def break_type(self) -> BreakType | None:
        """
        Gets/Sets the type of break that is applied at the beginning of the table.

        **optional**

        Returns:
            BreakType: Break Type

        Hint:
            - ``BreakType`` can be imported from ``ooo.dyn.style.break_type``
        """
        with contextlib.suppress(AttributeError):
            return self.__component.BreakType  # type: ignore
        return None

    @break_type.setter
    def break_type(self, value: BreakType) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.BreakType = value  # type: ignore

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
    def collapsing_borders(self) -> bool | None:
        """
        Gets/Sets whether borders of neighboring table cells are collapsed into one.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.CollapsingBorders
        return None

    @collapsing_borders.setter
    def collapsing_borders(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.CollapsingBorders = value

    @property
    def header_row_count(self) -> int | None:
        """
        Gets/Sets the number of rows of the table repeated on every new page.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.HeaderRowCount
        return None

    @header_row_count.setter
    def header_row_count(self, value: int) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.HeaderRowCount = value

    @property
    def hori_orient(self) -> HoriOrientationEnum:
        """
        Gets/Sets the horizontal orientation.

        When setting the value can be an integer or a ``HoriOrientationEnum``.

        Returns:
            HoriOrientationEnum: Horizontal Orientation

        Hint:
            - ``HoriOrientationEnum`` can be imported from ``ooo.dyn.text.hori_orientation``
        """
        # ooo.dyn.text.hori_orientation import HoriOrientationEnum
        return HoriOrientationEnum(self.__component.HoriOrient)

    @hori_orient.setter
    def hori_orient(self, value: int | HoriOrientationEnum) -> None:
        val = HoriOrientationEnum(value)
        self.__component.HoriOrient = val.value

    @property
    def is_width_relative(self) -> bool:
        """
        Gets/Sets if the value of the relative width is valid.
        """
        return self.__component.IsWidthRelative

    @is_width_relative.setter
    def is_width_relative(self, value: bool) -> None:
        self.__component.IsWidthRelative = value

    @property
    def keep_together(self) -> bool:
        """
        Gets/Sets - Setting this property to ``True`` prevents page or column breaks between this table and the following paragraph or text table.
        """
        return self.__component.KeepTogether

    @keep_together.setter
    def keep_together(self, value: bool) -> None:
        self.__component.KeepTogether = value

    @property
    def left_margin(self) -> int:
        """
        Gets/Sets the left margin of the table.
        """
        return self.__component.LeftMargin

    @left_margin.setter
    def left_margin(self, value: int) -> None:
        self.__component.LeftMargin = value

    @property
    def page_desc_name(self) -> str:
        """
        Gets/Sets - If this property is set, it creates a page break before the table and assigns the value as the name of the new page style sheet to use.
        """
        return self.__component.PageDescName

    @page_desc_name.setter
    def page_desc_name(self, value: str) -> None:
        self.__component.PageDescName = value

    @property
    def page_number_offset(self) -> int:
        """
        Gets/Sets - If a page break property is set at the table, this property contains the new value for the page number.
        """
        return self.__component.PageNumberOffset

    @page_number_offset.setter
    def page_number_offset(self, value: int) -> None:
        self.__component.PageNumberOffset = value

    @property
    def relative_width(self) -> int:
        """
        Gets/Sets the width of the table relative to its environment.
        """
        return self.__component.RelativeWidth

    @relative_width.setter
    def relative_width(self, value: int) -> None:
        self.__component.RelativeWidth = value

    @property
    def repeat_headline(self) -> bool:
        """
        Gets/Sets if the first row of the table is repeated on every new page.
        """
        return self.__component.RepeatHeadline

    @repeat_headline.setter
    def repeat_headline(self, value: bool) -> None:
        self.__component.RepeatHeadline = value

    @property
    def right_margin(self) -> int:
        """
        Gets/Sets the right margin of the table.
        """
        return self.__component.RightMargin

    @right_margin.setter
    def right_margin(self, value: int) -> None:
        self.__component.RightMargin = value

    @property
    def shadow_format(self) -> ShadowFormatStructComp:
        """
        Gets/Sets the type, color and size of the shadow.

        When setting the value can be an instance of ``ShadowFormatStructComp`` or ``ShadowFormat``.

        Returns:
            ShadowFormatStructComp: Shadow Format

        Hint:
            - ``ShadowFormat`` can be imported from ``ooo.dyn.table.shadow_format``
        """
        key = "ShadowFormat"
        prop = self.__props.get(key, None)
        if prop is None:
            prop = ShadowFormatStructComp(self.__component.ShadowFormat, key, self.__event_provider)
            self.__props[key] = prop
        return cast(ShadowFormatStructComp, prop)

    @shadow_format.setter
    def shadow_format(self, value: ShadowFormat | ShadowFormatStructComp) -> None:
        key = "ShadowFormat"
        if mInfo.Info.is_instance(value, ShadowFormatStructComp):
            self.__component.ShadowFormat = value.copy()
        else:
            self.__component.ShadowFormat = cast("ShadowFormat", value)
        if key in self.__props:
            del self.__props[key]

    @property
    def split(self) -> bool:
        """
        Get/Sets - Setting this property to ``False`` prevents the table from getting spread on two pages.
        """
        return self.__component.Split

    @split.setter
    def split(self, value: bool) -> None:
        self.__component.Split = value

    @property
    def table_border(self) -> TableBorderStructComp:
        """
        Gets/Sets a description of the cell or cell range border.

        If used with a cell range, the top, left, right, and bottom lines are at the edges of the entire range, not at the edges of the individual cell.

        Setting value can be done with a ``TableBorder`` or ``TableBorderStructComp`` object.

        Returns:
            TableBorderComp: Table Border.

        Hint:
            - ``TableBorder`` can be imported from ``ooo.dyn.table.table_border``
        """
        key = "TableBorder"
        prop = self.__props.get(key, None)
        if prop is None:
            prop = TableBorderStructComp(self.__component.TableBorder, key, self.__event_provider)
            self.__props[key] = prop
        return cast(TableBorderStructComp, prop)

    @table_border.setter
    def table_border(self, value: TableBorder | TableBorderStructComp) -> None:
        key = "TableBorder"
        if mInfo.Info.is_instance(value, TableBorderStructComp):
            self.__component.TableBorder = value.copy()
        else:
            self.__component.TableBorder = cast("TableBorder", value)
        if key in self.__props:
            del self.__props[key]

    @property
    def table_column_relative_sum(self) -> int:
        """
        Gets/Sets the sum of the column width values used in TableColumnSeparators.
        """
        return self.__component.TableColumnRelativeSum

    @table_column_relative_sum.setter
    def table_column_relative_sum(self, value: int) -> None:
        self.__component.TableColumnRelativeSum = value

    @property
    def table_template_name(self) -> str | None:
        """
        Gets/Sets the name of table style used by the table.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.TableTemplateName
        return None

    @table_template_name.setter
    def table_template_name(self, value: str) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.TableTemplateName = value

    @property
    def top_margin(self) -> int:
        """
        Gets/Sets the top margin.
        """
        return self.__component.TopMargin

    @top_margin.setter
    def top_margin(self, value: int) -> None:
        self.__component.TopMargin = value

    @property
    def width(self) -> int:
        """
        Gets/Sets the absolute table width.

        As this is only a describing property the value of the actual table may vary depending on the environment the table is located in and the settings of LeftMargin, RightMargin and HoriOrient.
        """
        return self.__component.Width

    @width.setter
    def width(self, value: int) -> None:
        self.__component.Width = value

    # endregion Properties
