from __future__ import annotations
from typing import cast, TYPE_CHECKING, Tuple
import contextlib
import uno

from ooodev.adapter.drawing.fill_properties_partial import FillPropertiesPartial
from ooodev.adapter.beans.property_set_partial import PropertySetPartial

if TYPE_CHECKING:
    from com.sun.star.awt import Gradient  # Struct
    from com.sun.star.awt import Size  # struct
    from com.sun.star.beans import XPropertySet
    from com.sun.star.chart2 import DataPointLabel
    from com.sun.star.chart2 import DataPointProperties
    from com.sun.star.chart2 import RelativePosition
    from com.sun.star.chart2 import Symbol  # struct
    from com.sun.star.chart2 import XDataPointCustomLabelField
    from com.sun.star.drawing import Hatch  # Struct
    from com.sun.star.drawing import LineDash  # Struct
    from com.sun.star.drawing.BitmapMode import BitmapModeProto  # type: ignore
    from com.sun.star.drawing.FillStyle import FillStyleProto  # type: ignore
    from com.sun.star.drawing.LineStyle import LineStyleProto  # type: ignore
    from com.sun.star.drawing.RectanglePoint import RectanglePointProto  # type: ignore
    from ooodev.utils.color import Color


class DataPointPropertiesPartial(PropertySetPartial, FillPropertiesPartial):
    """
    Partial class for XStyleFamiliesSupplier.

    See Also:
        `API DataPointProperties <https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1chart2_1_1DataPointProperties.html>`_
    """

    def __init__(self, component: DataPointProperties) -> None:
        """
        Constructor

        Args:
            component (DataPointProperties): UNO Component that implements ``com.sun.star.chart2.DataPointProperties`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``DataPointProperties``.
        """
        PropertySetPartial.__init__(self, component=component)
        FillPropertiesPartial.__init__(self, component=component)
        self.__component = component

    # region DataPointProperties
    @property
    def custom_label_fields(self) -> Tuple[XDataPointCustomLabelField, ...] | None:
        """
        Gets/Sets a text with possible fields that is used as a data point label, if set then Label property is ignored.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.CustomLabelFields
        return None

    @custom_label_fields.setter
    def custom_label_fields(self, value: Tuple[XDataPointCustomLabelField, ...]) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.CustomLabelFields = value

    @property
    def border_color(self) -> int:
        """
        Gets/Sets - Is used for borders around filled objects.

        See ``line_color``.
        """
        return self.__component.BorderColor

    @border_color.setter
    def border_color(self, value: int) -> None:
        self.__component.BorderColor = value

    @property
    def border_dash(self) -> LineDash:
        """
        Gets/Sets - Is used for borders around filled objects.

        See LineDash.
        """
        return self.__component.BorderDash

    @border_dash.setter
    def border_dash(self, value: LineDash) -> None:
        self.__component.BorderDash = value

    @property
    def border_dash_name(self) -> str | None:
        """
        Gets/Sets the name of a dash that can be found in the ``com.sun.star.container.XNameContainer``
        ``com.sun.star.drawing.LineDashTable``, that can be created via the ``com.sun.star.uno.XMultiServiceFactory`` of the Chart Document.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.BorderDashName
        return None

    @border_dash_name.setter
    def border_dash_name(self, value: str) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.BorderDashName = value

    @property
    def border_style(self) -> LineStyleProto:
        """
        Gets/Sets - Is used for borders around filled objects.

        See LineStyle.
        """
        return self.__component.BorderStyle

    @border_style.setter
    def border_style(self, value: LineStyleProto) -> None:
        self.__component.BorderStyle = value

    @property
    def border_transparency(self) -> int | None:
        """
        Gets/Sets - Is used for borders around filled objects.

        See ``line_transparence``.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.BorderTransparency

    @border_transparency.setter
    def border_transparency(self, value: int) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.BorderTransparency = value

    @property
    def border_width(self) -> int:
        """
        Gets/Sets Is used for borders around filled objects.

        See ``line_width``.
        """
        return self.__component.BorderWidth

    @border_width.setter
    def border_width(self, value: int) -> None:
        self.__component.BorderWidth = value

    @property
    def color(self) -> Color:
        """
        Gets/Sets - points to a style that also supports this service (but not this property) that is used as default,
        if the ``PropertyState`` of a property is ``DEFAULT_VALUE``.

        This is the main color of a data point.

        For charts with filled areas, like bar-charts, this should map to the ``fill_color`` of the objects.
        For line-charts this should map to the ``line_color`` property.
        """
        return Color(self.__component.Color)

    @color.setter
    def color(self, value: Color) -> None:
        self.__component.Color = cast(int, value)

    @property
    def custom_label_position(self) -> RelativePosition | None:
        """
        Gets/Sets - Custom position on the page associated to the CUSTOM label placement.

        **since**

            LibreOffice ``7.0``

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.CustomLabelPosition
        return None

    @custom_label_position.setter
    def custom_label_position(self, value: RelativePosition) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.CustomLabelPosition = value

    @property
    def error_bar_x(self) -> XPropertySet | None:
        """
        Gets/Sets - If void, no error bars are shown for the data point in x-direction.

        The ``com.sun.star.beans.XPropertySet`` must support the service ``ErrorBar``.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.ErrorBarX
        return None

    @error_bar_x.setter
    def error_bar_x(self, value: XPropertySet) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.ErrorBarX = value

    @property
    def error_bar_y(self) -> XPropertySet | None:
        """
        Get/Sets - If void, no error bars are shown for the data point in y-direction.

        The ``com.sun.star.beans.XPropertySet`` must support the service ErrorBar.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.ErrorBarY
        return None

    @error_bar_y.setter
    def error_bar_y(self, value: XPropertySet) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.ErrorBarY = value

    @property
    def fill_background(self) -> bool:
        """
        Gets/Sets - If ``True``, fills the background of a hatch with the color given in the Color property.
        """
        return self.__component.FillBackground

    @fill_background.setter
    def fill_background(self, value: bool) -> None:
        self.__component.FillBackground = value

    @property
    def fill_bitmap_logical_size(self) -> bool:
        """
        Gets/Sets if the size is given in percentage or as an absolute value.

        If this is ``True``, the properties ``fill_bitmap_size_x`` and ``fill_bitmap_size_y`` contain the size of the
        tile in percent of the size of the original bitmap. If this is ``False``, the size of the tile is specified with ``1/100th mm``.
        """
        return self.__component.FillBitmapLogicalSize

    @fill_bitmap_logical_size.setter
    def fill_bitmap_logical_size(self, value: bool) -> None:
        self.__component.FillBitmapLogicalSize = value

    @property
    def fill_bitmap_mode(self) -> BitmapModeProto:
        """
        Gets/Sets - this enum selects how an area is filled with a single bitmap.
        """
        return self.__component.FillBitmapMode

    @fill_bitmap_mode.setter
    def fill_bitmap_mode(self, value: BitmapModeProto) -> None:
        self.__component.FillBitmapMode = value

    @property
    def fill_bitmap_name(self) -> str:
        """
        Gets/Sets - The name of the fill bitmap.
        """
        return self.__component.FillBitmapName

    @fill_bitmap_name.setter
    def fill_bitmap_name(self, value: str) -> None:
        self.__component.FillBitmapName = value

    @property
    def fill_bitmap_offset_x(self) -> int:
        """
        Gets/Sets - This is the horizontal offset where the tile starts.

        It is given in percent in relation to the width of the bitmap.
        """
        return self.__component.FillBitmapOffsetX

    @fill_bitmap_offset_x.setter
    def fill_bitmap_offset_x(self, value: int) -> None:
        self.__component.FillBitmapOffsetX = value

    @property
    def fill_bitmap_offset_y(self) -> int:
        """
        Gets/Sets - This is the vertical offset where the tile starts.

        It is given in percent in relation to the width of the bitmap.
        """
        return self.__component.FillBitmapOffsetY

    @fill_bitmap_offset_y.setter
    def fill_bitmap_offset_y(self, value: int) -> None:
        self.__component.FillBitmapOffsetY = value

    @property
    def fill_bitmap_position_offset_x(self) -> int:
        """
        Gets/Sets - Every second line of tiles is moved the given percent of the width of the bitmap.
        """
        return self.__component.FillBitmapPositionOffsetX

    @fill_bitmap_position_offset_x.setter
    def fill_bitmap_position_offset_x(self, value: int) -> None:
        self.__component.FillBitmapPositionOffsetX = value

    @property
    def fill_bitmap_position_offset_y(self) -> int:
        """
        Gets/Sets - Every second row of tiles is moved the given percent of the width of the bitmap.
        """
        return self.__component.FillBitmapPositionOffsetY

    @fill_bitmap_position_offset_y.setter
    def fill_bitmap_position_offset_y(self, value: int) -> None:
        self.__component.FillBitmapPositionOffsetY = value

    @property
    def fill_bitmap_rectangle_point(self) -> RectanglePointProto:
        """
        Gets/Sets - The RectanglePoint specifies the position inside of the bitmap to use as the top left position for rendering.
        """
        return self.__component.FillBitmapRectanglePoint

    @fill_bitmap_rectangle_point.setter
    def fill_bitmap_rectangle_point(self, value: RectanglePointProto) -> None:
        self.__component.FillBitmapRectanglePoint = value

    @property
    def fill_bitmap_size_x(self) -> int:
        """
        Gets/Sets - This is the width of the tile for filling.

        Depending on the property FillBitmapLogicalSize, this is either relative or absolute.
        """
        return self.__component.FillBitmapSizeX

    @fill_bitmap_size_x.setter
    def fill_bitmap_size_x(self, value: int) -> None:
        self.__component.FillBitmapSizeX = value

    @property
    def fill_bitmap_size_y(self) -> int:
        """
        Gets/Sets - This is the height of the tile for filling.

        Depending on the property FillBitmapLogicalSize, this is either relative or absolute.
        """
        return self.__component.FillBitmapSizeY

    @fill_bitmap_size_y.setter
    def fill_bitmap_size_y(self, value: int) -> None:
        self.__component.FillBitmapSizeY = value

    @property
    def fill_style(self) -> FillStyleProto:
        """
        Gets/Sets - This enumeration selects the style with which the area will be filled.
        """
        return self.__component.FillStyle

    @fill_style.setter
    def fill_style(self, value: FillStyleProto) -> None:
        self.__component.FillStyle = value

    @property
    def geometry3d(self) -> int | None:
        """
        Gets/Sets the geometry of a 3 dimensional data point.

        Number is one of constant group DataPointGeometry3D.

        This is especially used for 3D bar-charts.

        ``CUBOID==0 CYLINDER==1 CONE==2 PYRAMID==3 CUBOID==else``

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.Geometry3D
        return None

    @geometry3d.setter
    def geometry3d(self, value: int) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.Geometry3D = value

    @property
    def gradient(self) -> Gradient | None:
        """
        Gets/Sets the gradient.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.Gradient
        return None

    @gradient.setter
    def gradient(self, value: Gradient) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.Gradient = value

    @property
    def gradient_name(self) -> str:
        """
        Gets/Sets the name of the gradient.
        """
        return self.__component.GradientName

    @gradient_name.setter
    def gradient_name(self, value: str) -> None:
        self.__component.GradientName = value

    @property
    def hatch(self) -> Hatch | None:
        """
        Gets/Sets the hatch.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.Hatch
        return None

    @hatch.setter
    def hatch(self, value: Hatch) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.Hatch = value

    @property
    def hatch_name(self) -> str:
        """
        Gets/Sets the name of the hatch.
        """
        return self.__component.HatchName

    @hatch_name.setter
    def hatch_name(self, value: str) -> None:
        self.__component.HatchName = value

    @property
    def label(self) -> DataPointLabel:
        """
        Gets/Sets the data point label.
        """
        return self.__component.Label

    @label.setter
    def label(self, value: DataPointLabel) -> None:
        self.__component.Label = value

    @property
    def label_placement(self) -> int | None:
        """
        Gets/Sets a relative position for the data label.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.LabelPlacement
        return None

    @label_placement.setter
    def label_placement(self, value: int) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.LabelPlacement = value

    @property
    def label_separator(self) -> str | None:
        """
        Gets/Sets a string that is used to separate the parts of a data label (caption).

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.LabelSeparator
        return None

    @label_separator.setter
    def label_separator(self, value: str) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.LabelSeparator = value

    @property
    def line_dash(self) -> LineDash:
        """
        Gets/Sets - Is only used for line-chart types.
        """
        return self.__component.LineDash

    @line_dash.setter
    def line_dash(self, value: LineDash) -> None:
        self.__component.LineDash = value

    @property
    def line_dash_name(self) -> str | None:
        """
        Gets/Sets the name of a dash that can be found in the ``com.sun.star.container.XNameContainer``
        ``com.sun.star.drawing.LineDashTable``, that can be created via the ``com.sun.star.uno.XMultiServiceFactory`` of the Chart Document.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.LineDashName
        return None

    @line_dash_name.setter
    def line_dash_name(self, value: str) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.LineDashName = value

    @property
    def line_style(self) -> LineStyleProto:
        """
        Gets/Sets the line style.
        """
        return self.__component.LineStyle

    @line_style.setter
    def line_style(self, value: LineStyleProto) -> None:
        self.__component.LineStyle = value

    @property
    def line_width(self) -> int:
        """
        Gets/Sets - Is only used for line-chart types.
        """
        return self.__component.LineWidth

    @line_width.setter
    def line_width(self, value: int) -> None:
        self.__component.LineWidth = value

    @property
    def number_format(self) -> int | None:
        """
        Gets/Sets a number format for the display of the value in the data label.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.NumberFormat
        return None

    @number_format.setter
    def number_format(self, value: int) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.NumberFormat = value

    @property
    def offset(self) -> float | None:
        """
        Gets/sets a value by which a data point is moved from its default position in percent of the maximum allowed distance.

        This is especially useful for the explosion of pie-chart segments.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.Offset
        return None

    @offset.setter
    def offset(self, value: float) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.Offset = value

    @property
    def percent_diagonal(self) -> int | None:
        """
        Gets/Seta a value between ``0`` and ``100`` indicating the percentage how round an edge should be.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.PercentDiagonal
        return None

    @percent_diagonal.setter
    def percent_diagonal(self, value: int) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.PercentDiagonal = value

    @property
    def percentage_number_format(self) -> int | None:
        """
        Gets/Sets a number format for the display of the percentage value in the data label.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.PercentageNumberFormat
        return None

    @percentage_number_format.setter
    def percentage_number_format(self, value: int) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.PercentageNumberFormat = value

    @property
    def reference_page_size(self) -> Size | None:
        """
        Gets/Sets - The size of the page at the moment when the font size for data labels was set.

        This size is used to resize text in the view when the size of the page has changed since the font sizes were set (automatic text scaling).

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.ReferencePageSize
        return None

    @reference_page_size.setter
    def reference_page_size(self, value: Size) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.ReferencePageSize = value

    @property
    def show_error_box(self) -> bool | None:
        """
        Gets/Sets - In case ``error_bar_x`` and ``error_bar_y`` both are set, and error bars are shown,
        a box spanning all error-indicators is rendered.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.ShowErrorBox
        return None

    @show_error_box.setter
    def show_error_box(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.ShowErrorBox = value

    @property
    def symbol(self) -> Symbol | None:
        """
        Gets/Sets - This is the symbol that is used for a data point.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.Symbol
        return None

    @symbol.setter
    def symbol(self, value: Symbol) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.Symbol = value

    @property
    def text_word_wrap(self) -> bool | None:
        """
        Gets/Sets if the text of a data label (caption) must be wrapped

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.TextWordWrap
        return None

    @text_word_wrap.setter
    def text_word_wrap(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.TextWordWrap = value

    @property
    def transparency(self) -> int:
        """
        Gets/Sets - This is the main transparency value of a data point.

        For charts with filled areas, like bar-charts, this should map to the FillTransparence of the objects.
        For line-charts this should map to the ``line_transparence`` property.
        """
        return self.__component.Transparency

    @transparency.setter
    def transparency(self, value: int) -> None:
        self.__component.Transparency = value

    @property
    def transparency_gradient(self) -> Gradient | None:
        """
        Gets/Sets the transparency of the fill area as a gradient.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.TransparencyGradient
        return None

    @transparency_gradient.setter
    def transparency_gradient(self, value: Gradient) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.TransparencyGradient = value

    @property
    def transparency_gradient_name(self) -> str:
        """
        Gets/Sets the name of the transparency gradient.
        """
        return self.__component.TransparencyGradientName

    @transparency_gradient_name.setter
    def transparency_gradient_name(self, value: str) -> None:
        self.__component.TransparencyGradientName = value
        # endregion DataPointProperties
