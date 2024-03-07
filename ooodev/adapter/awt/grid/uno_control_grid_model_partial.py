from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING, Tuple
import uno  # pylint: disable=unused-import
from ooo.dyn.text.font_emphasis import FontEmphasisEnum
from ooo.dyn.text.font_relief import FontReliefEnum
from ooo.dyn.style.vertical_alignment import VerticalAlignment
from ooo.dyn.view.selection_type import SelectionType
from ooodev.utils import info as mInfo
from ooodev.events.events import Events
from ooodev.utils.color import Color
from ooodev.utils.partial.model_prop_partial import ModelPropPartial
from ooodev.adapter.awt.uno_control_model_partial import UnoControlModelPartial
from ooodev.adapter.awt.font_descriptor_struct_comp import FontDescriptorStructComp

if TYPE_CHECKING:
    from com.sun.star.awt.grid import UnoControlGridModel  # Service
    from com.sun.star.awt import FontDescriptor  # struct
    from com.sun.star.awt.grid import XGridColumnModel
    from ooodev.events.args.key_val_args import KeyValArgs


class UnoControlGridModelPartial(UnoControlModelPartial):
    """Partial class for UnoControlGridModel."""

    def __init__(self):
        """
        Constructor

        Args:
            component (Any): Component that implements ``com.sun.star.awt.UnoControlGridModel`` service.
        """
        # pylint: disable=unused-argument
        if not isinstance(self, ModelPropPartial):
            raise TypeError("This class must be used as a mixin that implements ModelPropPartial.")

        self.model: UnoControlGridModel
        UnoControlModelPartial.__init__(self)
        self.__event_provider = Events(self)
        self.__props = {}

        def on_comp_struct_changed(src: Any, event_args: KeyValArgs) -> None:
            prop_name = str(event_args.event_data["prop_name"])
            if hasattr(self.model, prop_name):
                setattr(self.model, prop_name, event_args.source.component)

        self.__fn_on_comp_struct_changed = on_comp_struct_changed

        self.__event_provider.subscribe_event(
            "com_sun_star_awt_FontDescriptor_changed", self.__fn_on_comp_struct_changed
        )

    def set_font_descriptor(self, font_descriptor: FontDescriptor | FontDescriptorStructComp) -> None:
        """
        Sets the font descriptor of the control.

        Args:
            font_descriptor (FontDescriptor, FontDescriptorStructComp): UNO Struct - Font descriptor to set.

        Note:
            The ``font_descriptor`` property can also be used to set the font descriptor.

        Hint:
            - ``FontDescriptor`` can be imported from ``ooo.dyn.awt.font_descriptor``.
        """
        self.font_descriptor = font_descriptor

    # region Properties

    @property
    def font_descriptor(self) -> FontDescriptorStructComp:
        """
        Gets/Sets the Font Descriptor.

        Setting value can be done with a ``FontDescriptor`` or ``FontDescriptorStructComp`` object.

        Returns:
            ~ooodev.adapter.awt.font_descriptor_struct_comp.FontDescriptorStructComp: Font Descriptor

        Hint:
            - ``FontDescriptor`` can be imported from ``ooo.dyn.awt.font_descriptor``.
        """
        key = "FontDescriptor"
        prop = self.__props.get(key, None)
        if prop is None:
            prop = FontDescriptorStructComp(self.model.FontDescriptor, key, self.__event_provider)
            self.__props[key] = prop
        return cast(FontDescriptorStructComp, prop)

    @font_descriptor.setter
    def font_descriptor(self, value: FontDescriptor | FontDescriptorStructComp) -> None:
        key = "FontDescriptor"
        if mInfo.Info.is_instance(value, FontDescriptorStructComp):
            self.model.FontDescriptor = value.copy()
        else:
            self.model.FontDescriptor = cast("FontDescriptor", value)
        if key in self.__props:
            del self.__props[key]

    @property
    def row_background_colors(self) -> Tuple[Color, ...]:
        """
        Gets/Sets the colors to be used as background for data rows.

        If this sequence is non-empty, the data rows will be rendered with alternating background colors: Assuming the sequence has n elements, each row will use the background color as specified by its number's remainder modulo n.

        If this sequence is empty, all rows will use the same background color as the control as whole.

        If this property is empty, rows will be painted in alternating background colors, every second row having a background color derived from the control's selection color.

        Returns:
            Tuple[~ooodev.utils.color.Color, ...]: Tuple of ~ooodev.utils.color.Color objects.
        """
        result = self.model.RowBackgroundColors
        return result or ()  # type: ignore

    @row_background_colors.setter
    def row_background_colors(self, value: Tuple[Color, ...]) -> None:
        self.model.RowBackgroundColors = value  # type: ignore

    @property
    def active_selection_background_color(self) -> Color | None:
        """
        Gets/Sets the color to be used when drawing the background of selected cells, while the control has the focus.

        If this property has a value of ``None``, the grid control renderer will use some default color, depending on the control's style settings.

        Returns:
            ~ooodev.utils.color.Color | None: Color to be used when drawing the background of selected cells, while the control has the focus.
        """
        result = self.model.ActiveSelectionBackgroundColor
        return None if result is None else Color(result)

    @active_selection_background_color.setter
    def active_selection_background_color(self, value: Color) -> None:
        self.model.ActiveSelectionBackgroundColor = value  # type: ignore

    @property
    def active_selection_text_color(self) -> Color | None:
        """
        Gets the color to be used when drawing the text of selected cells, while the control has the focus.

        If this property has a value of VOID, the grid control renderer will use some default color, depending on the control's style settings.

        Returns:
            ~ooodev.utils.color.Color | None: Color or None if not set.
        """
        result = self.model.ActiveSelectionTextColor
        return None if result is None else Color(result)

    @active_selection_text_color.setter
    def active_selection_text_color(self, value: Color) -> None:
        self.model.ActiveSelectionTextColor = value  # type: ignore

    @property
    def column_header_height(self) -> int | None:
        """
        Gets/Sets the height of the column header row, if applicable.

        The height is specified in application font units - see ``MeasureUnit`` below.

        The value given here is ignored if ShowColumnHeader is FALSE.

        If the property is ``None``, the grid control shall automatically determine a height which conveniently allows,
        according to the used font, to display one line of text.

        See Also:
            - `com.sun.star.util.MeasureUnit <https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1util_1_1MeasureUnit.html>`__
        """
        return self.model.ColumnHeaderHeight

    @column_header_height.setter
    def column_header_height(self, value: int) -> None:
        self.model.ColumnHeaderHeight = value

    @property
    def column_model(self) -> XGridColumnModel:
        """
        Specifies the ``XGridColumnModel`` that is providing the column structure.

        You can implement your own instance of ``XGridColumnModel`` or use the DefaultGridColumnModel.

        The column model is in the ownership of the grid model: When you set a new column model, or dispose the grid model, then the (old) column model is disposed, too.

        The default for this property is an empty instance of the DefaultGridColumnModel.
        """
        return self.model.ColumnModel

    @column_model.setter
    def column_model(self, value: XGridColumnModel) -> None:
        self.model.ColumnModel = value

    @property
    def font_emphasis_mark(self) -> FontEmphasisEnum:
        """
        Gets/Sets the ``FontEmphasis`` value of the text in the control.

        Note:
            Value can be set with ``FontEmphasisEnum`` or ``int``.

        Hint:
            - ``FontEmphasisEnum`` can be imported from ``ooo.dyn.text.font_emphasis``.
        """
        return FontEmphasisEnum(self.model.FontEmphasisMark)

    @font_emphasis_mark.setter
    def font_emphasis_mark(self, value: int | FontEmphasisEnum) -> None:
        self.model.FontEmphasisMark = int(value)

    @property
    def font_relief(self) -> FontReliefEnum:
        """
        Gets/Sets ``FontRelief`` value of the text in the control.

        Note:
            Value can be set with ``FontReliefEnum`` or ``int``.

        Hint:
            - ``FontReliefEnum`` can be imported from ``ooo.dyn.text.font_relief``.
        """
        return FontReliefEnum(self.model.FontRelief)

    @font_relief.setter
    def font_relief(self, value: int | FontReliefEnum) -> None:
        self.model.FontRelief = int(value)

    @property
    def grid_line_color(self) -> Color | None:
        """
        Gets/Sets the color to be used when drawing lines between cells

        If this property has a value of ``None``, the grid control renderer will use some default color,
        depending on the control's style settings.

        Returns:
            ~ooodev.utils.color.Color | None: Color or None if not set.
        """
        result = self.model.GridLineColor
        return None if result is None else Color(result)

    @grid_line_color.setter
    def grid_line_color(self, value: Color) -> None:
        self.model.GridLineColor = value  # type: ignore

    @property
    def h_scroll(self) -> bool:
        """
        Gets/Sets the vertical scrollbar mode.

        The default value is ``False``
        """
        return self.model.HScroll

    @h_scroll.setter
    def h_scroll(self, value: bool) -> None:
        self.model.HScroll = value

    @property
    def header_background_color(self) -> Color | None:
        """
        Gets/Sets the color to be used when drawing the background of row or column headers

        If this property has a value of ``None``, the grid control renderer will use some default color,
        depending on the control's style settings.

        Returns:
            ~ooodev.utils.color.Color | None: Color or None if not set.
        """
        result = self.model.HeaderBackgroundColor
        return None if result is None else Color(result)

    @header_background_color.setter
    def header_background_color(self, value: Color) -> None:
        self.model.HeaderBackgroundColor = value  # type: ignore

    @property
    def header_text_color(self) -> Color | None:
        """
        Gets/Sets the color to be used when drawing the text within row or column headers

        If this property has a value of ``None``, the grid control renderer will use some default color,
        depending on the control's style settings.

        Returns:
            ~ooodev.utils.color.Color | None: Color or None if not set.
        """
        result = self.model.HeaderTextColor
        return None if result is None else Color(result)

    @header_text_color.setter
    def header_text_color(self, value: Color) -> None:
        self.model.HeaderTextColor = value  # type: ignore

    @property
    def help_text(self) -> str:
        """
        Get/Sets the help text of the control.
        """
        return self.model.HelpText

    @help_text.setter
    def help_text(self, value: str) -> None:
        self.model.HelpText = value

    @property
    def help_url(self) -> str:
        """
        Gets/Sets the help URL of the control.
        """
        return self.model.HelpURL

    @help_url.setter
    def help_url(self, value: str) -> None:
        self.model.HelpURL = value

    @property
    def inactive_selection_background_color(self) -> Color | None:
        """
        Gets/Sets the color to be used when drawing the background of selected cells, while the control does not have the focus.

        If this property has a value of ``None``, the grid control renderer will use some default color,
        depending on the control's style settings.

        Returns:
            ~ooodev.utils.color.Color | None: Color or None if not set.
        """
        result = self.model.InactiveSelectionBackgroundColor
        return None if result is None else Color(result)

    @inactive_selection_background_color.setter
    def inactive_selection_background_color(self, value: Color) -> None:
        self.model.InactiveSelectionBackgroundColor = value  # type: ignore

    @property
    def inactive_selection_text_color(self) -> Color | None:
        """
        Gets/Sets the color to be used when drawing the text of selected cells, while the control does not have the focus.

        If this property has a value of ``None``, the grid control renderer will use some default color,
        depending on the control's style settings.

        Returns:
            ~ooodev.utils.color.Color | None: Color or None if not set.
        """
        result = self.model.InactiveSelectionTextColor
        return None if result is None else Color(result)

    @inactive_selection_text_color.setter
    def inactive_selection_text_color(self, value: Color) -> None:
        self.model.InactiveSelectionTextColor = value  # type: ignore

    @property
    def row_header_width(self) -> int:
        """
        Gets/Sets the width of the row header column, if applicable.

        The width is specified in application font units - see ``MeasureUnit`` below.

        The value given here is ignored if ``show_row_header`` is ``False``.

        See Also:
            - `com.sun.star.util.MeasureUnit <https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1util_1_1MeasureUnit.html>`__
        """
        return self.model.RowHeaderWidth

    @row_header_width.setter
    def row_header_width(self, value: int) -> None:
        self.model.RowHeaderWidth = value

    @property
    def row_height(self) -> int:
        """
        Gets/Sets the height of rows in the grid control.

        The height is specified in application font units - see ``MeasureUnit`` below.

        See Also:
            - `com.sun.star.util.MeasureUnit <https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1util_1_1MeasureUnit.html>`__
        """
        return self.model.RowHeight

    @row_height.setter
    def row_height(self, value: int) -> None:
        self.model.RowHeight = value

    @property
    def selection_model(self) -> SelectionType:
        """
        Gets/Sets the selection mode that is enabled for this grid control.

        The default value is ``com.sun.star.view.SelectionType.SINGLE``

        Hint:
            - ``SelectionType`` can be imported from ``ooo.dyn.view.selection_type``
        """
        return self.model.SelectionModel  # type: ignore

    @selection_model.setter
    def selection_model(self, value: SelectionType) -> None:
        self.model.SelectionModel = value  # type: ignore

    @property
    def show_column_header(self) -> bool:
        """
        Gets/Sets whether the grid control should display a title row.

        The default value is ``True``.
        """
        return self.model.ShowColumnHeader

    @show_column_header.setter
    def show_column_header(self, value: bool) -> None:
        self.model.ShowColumnHeader = value

    @property
    def show_row_header(self) -> bool:
        """
        Gets/Sets whether the grid control should display a special header column.

        The default value is ``False``.
        """
        return self.model.ShowRowHeader

    @property
    def tabstop(self) -> bool:
        """
        Gets/Sets - that the control can be reached with the TAB key.
        """
        return self.model.Tabstop

    @tabstop.setter
    def tabstop(self, value: bool) -> None:
        self.model.Tabstop = value

    @show_row_header.setter
    def show_row_header(self, value: bool) -> None:
        self.model.ShowRowHeader = value

    @property
    def text_color(self) -> Color:
        """
        Gets/Sets the text color of the control.

        Returns:
            ~ooodev.utils.color.Color: Text color.
        """
        return Color(self.model.TextColor)

    @text_color.setter
    def text_color(self, value: Color) -> None:
        self.model.TextColor = value  # type: ignore

    @property
    def text_line_color(self) -> Color:
        """
        Gets/Sets the text line color of the control.

        Returns:
            ~ooodev.utils.color.Color: Text line color.
        """
        return Color(self.model.TextLineColor)

    @text_line_color.setter
    def text_line_color(self, value: Color) -> None:
        self.model.TextLineColor = value  # type: ignore

    @property
    def use_grid_lines(self) -> bool:
        """
        Gets/Sets whether or not to paint horizontal and vertical lines between the grid cells.
        """
        return self.model.UseGridLines

    @use_grid_lines.setter
    def use_grid_lines(self, value: bool) -> None:
        self.model.UseGridLines = value

    @property
    def v_scroll(self) -> bool:
        """
        Gets/Sets the horizontal scrollbar mode.

        The default value is ``False``.
        """
        return self.model.VScroll

    @v_scroll.setter
    def v_scroll(self, value: bool) -> None:
        self.model.VScroll = value

    @property
    def vertical_align(self) -> VerticalAlignment:
        """
        Gets/Sets the vertical alignment of the text in the control.

        **optional**

        Hint:
            - ``VerticalAlignment`` can be imported from ``ooo.dyn.style.vertical_alignment``
        """
        return self.model.VerticalAlign  # type: ignore

    @vertical_align.setter
    def vertical_align(self, value: VerticalAlignment) -> None:
        self.model.VerticalAlign = value  # type: ignore

    # endregion Properties
