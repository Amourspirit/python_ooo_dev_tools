from __future__ import annotations
import contextlib
from typing import Any, TYPE_CHECKING
from ooo.dyn.text.font_emphasis import FontEmphasisEnum
from ooo.dyn.text.font_relief import FontReliefEnum
from ooo.dyn.style.vertical_alignment import VerticalAlignment
from ooo.dyn.awt.line_end_format import LineEndFormatEnum
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.utils.kind.align_kind import AlignKind
from ooodev.utils.kind.border_kind import BorderKind
from ooodev.utils.color import Color
from .uno_control_model_partial import UnoControlModelPartial
from .font_descriptor_comp import FontDescriptorComp

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlEditModel  # Service
    from com.sun.star.awt import FontDescriptor  # struct
    from ooodev.events.args.key_val_args import KeyValArgs


class UnoControlEditModelPartial(UnoControlModelPartial):
    """Partial class for UnoControlEditModel."""

    def __init__(self, component: UnoControlEditModel):
        """
        Constructor

        Args:
            component (Any): Component that implements ``com.sun.star.awt.UnoControlEditModel`` service.
        """
        # pylint: disable=unused-argument
        self.__component = component
        event_provider = self if isinstance(self, EventsPartial) else None
        UnoControlModelPartial.__init__(self, component=component)
        self.__font_descriptor = FontDescriptorComp(self.__component.FontDescriptor, event_provider)

        if event_provider is not None:

            def on_font_descriptor_changed(src: Any, event_args: KeyValArgs) -> None:
                self.__component.FontDescriptor = self.__font_descriptor.component

            self.__fn_on_font_descriptor_changed = on_font_descriptor_changed
            # pylint: disable=no-member
            event_provider.subscribe_event("font_descriptor_struct_changed", self.__fn_on_font_descriptor_changed)

    def set_font_descriptor(self, font_descriptor: FontDescriptor) -> None:
        """
        Sets the font descriptor of the control.

        Args:
            font_descriptor (FontDescriptor): UNO Struct - Font descriptor to set.
        """
        # FontDescriptorComp do not have any state, so we can directly assign the component.
        self.__font_descriptor.component = font_descriptor

    # region Properties

    @property
    def font_descriptor(self) -> FontDescriptorComp:
        """
        Gets the Font Descriptor

        Hint:
            ``set_font_descriptor()`` can be used to set the font descriptor.
        """
        return self.__font_descriptor

    @property
    def align(self) -> AlignKind:
        """
        Get/Sets the horizontal alignment of the text in the control.

        Hint:
            - ``AlignKind`` can be imported from ``ooodev.utils.kind.align_kind``.
        """
        return AlignKind(self.__component.Align)

    @align.setter
    def align(self, value: AlignKind | int) -> None:
        kind = AlignKind(int(value))
        self.__component.Align = kind.value

    @property
    def auto_h_scroll(self) -> bool:
        """
        Gets/Sets - If set to ``True`` an horizontal scrollbar will be added automatically when needed.
        """
        return self.__component.AutoHScroll

    @auto_h_scroll.setter
    def auto_h_scroll(self, value: bool) -> None:
        self.__component.AutoHScroll = value

    @property
    def auto_v_scroll(self) -> bool:
        """
        Gets/Sets - If set to ``True`` a vertical scrollbar will be added automatically when needed.
        """
        return self.__component.AutoVScroll

    @auto_v_scroll.setter
    def auto_v_scroll(self, value: bool) -> None:
        self.__component.AutoVScroll = value

    @property
    def background_color(self) -> Color:
        """
        Gets/Set the background color of the control.
        """
        return Color(self.__component.BackgroundColor)

    @background_color.setter
    def background_color(self, value: Color) -> None:
        self.__component.BackgroundColor = value  # type: ignore

    @property
    def border(self) -> BorderKind:
        """
        Gets/Sets the border style of the control.

        Note:
            Value can be set with ``BorderKind`` or ``int``.

        Hint:
            - ``BorderKind`` can be imported from ``ooodev.utils.kind.border_kind``.
        """
        return BorderKind(self.__component.Border)

    @border.setter
    def border(self, value: int | BorderKind) -> None:
        kind = BorderKind(int(value))
        self.__component.Border = kind.value

    @property
    def border_color(self) -> Color:
        """
        Gets/Sets the color of the border, if present

        Not every border style (see Border) may support coloring. For instance, usually a border with 3D effect will ignore the border_color setting.
        """
        return Color(self.__component.BorderColor)

    @border_color.setter
    def border_color(self, value: Color) -> None:
        self.__component.BorderColor = value

    @property
    def echo_char(self) -> str:
        """Gets/Sets the echo character as a string"""
        with contextlib.suppress(Exception):
            return chr(self.__component.EchoChar)
        return ""

    @echo_char.setter
    def echo_char(self, value: str) -> None:
        if value != "":
            value = value[0]  # first char
        self.__component.EchoChar = ord(value)
        ...

    @property
    def enabled(self) -> bool:
        """
        Gets/Sets whether the control is enabled or disabled.
        """
        return self.__component.Enabled

    @enabled.setter
    def enabled(self, value: bool) -> None:
        self.__component.Enabled = value

    @property
    def font_emphasis_mark(self) -> FontEmphasisEnum:
        """
        Gets/Sets the ``FontEmphasis`` value of the text in the control.

        Note:
            Value can be set with ``FontEmphasisEnum`` or ``int``.

        Hint:
            - ``FontEmphasisEnum`` can be imported from ``ooo.dyn.text.font_emphasis``.
        """
        return FontEmphasisEnum(self.__component.FontEmphasisMark)

    @font_emphasis_mark.setter
    def font_emphasis_mark(self, value: int | FontEmphasisEnum) -> None:
        self.__component.FontEmphasisMark = int(value)

    @property
    def font_relief(self) -> FontReliefEnum:
        """
        Gets/Sets ``FontRelief`` value of the text in the control.

        Note:
            Value can be set with ``FontReliefEnum`` or ``int``.

        Hint:
            - ``FontReliefEnum`` can be imported from ``ooo.dyn.text.font_relief``.
        """
        return FontReliefEnum(self.__component.FontRelief)

    @font_relief.setter
    def font_relief(self, value: int | FontReliefEnum) -> None:
        self.__component.FontRelief = int(value)

    @property
    def h_scroll(self) -> bool:
        """
        Gets/Sets if the content of the control can be scrolled in the horizontal direction.
        """
        return self.__component.HScroll

    @h_scroll.setter
    def h_scroll(self, value: bool) -> None:
        self.__component.HScroll = value

    @property
    def hard_line_breaks(self) -> bool:
        """
        Gets/Sets if hard line breaks will be returned in the ``XTextComponent.getText()`` method.
        """
        return self.__component.HardLineBreaks

    @hard_line_breaks.setter
    def hard_line_breaks(self, value: bool) -> None:
        self.__component.HardLineBreaks = value

    @property
    def help_text(self) -> str:
        """
        Get/Sets the help text of the control.
        """
        return self.__component.HelpText

    @help_text.setter
    def help_text(self, value: str) -> None:
        self.__component.HelpText = value

    @property
    def help_url(self) -> str:
        """
        Gets/Sets the help URL of the control.
        """
        return self.__component.HelpURL

    @help_url.setter
    def help_url(self, value: str) -> None:
        self.__component.HelpURL = value

    @property
    def hide_inactive_selection(self) -> bool:
        """
        Gets/Sets whether the selection in the control should be hidden when the control is not active (focused).
        """
        return self.__component.HideInactiveSelection

    @hide_inactive_selection.setter
    def hide_inactive_selection(self, value: bool) -> None:
        self.__component.HideInactiveSelection = value

    @property
    def line_end_format(self) -> LineEndFormatEnum:
        """
        specifies which line end type should be used for multi line text

        Controls working with this model care for this setting when the user enters text. Every line break entered into the control will be treated according to this setting, so that the Text property always contains only line ends in the format specified.

        Possible values are all constants from the LineEndFormat group.

        Note that this setting is usually not relevant when you set new text via the API. No matter which line end format is used in this new text then, usual control implementations should recognize all line end formats and display them properly.

        Note:
            Value can be set with ``LineEndFormatEnum`` or ``int``.

        Hint:
            - ``LineEndFormatEnum`` can be imported from ``ooo.dyn.awt.line_end_format``.
        """
        return LineEndFormatEnum(self.__component.LineEndFormat)

    @line_end_format.setter
    def line_end_format(self, value: int | LineEndFormatEnum) -> None:
        self.__component.LineEndFormat = int(value)

    @property
    def max_text_len(self) -> int:
        """
        Gets/Sets the maximum character count.

        There's no limitation, if set to 0.
        """
        return self.__component.MaxTextLen

    @max_text_len.setter
    def max_text_len(self, value: int) -> None:
        self.__component.MaxTextLen = value

    @property
    def multi_line(self) -> bool:
        """
        Gets/Sets that the text may be displayed on more than one line.
        """
        return self.__component.MultiLine

    @multi_line.setter
    def multi_line(self, value: bool) -> None:
        self.__component.MultiLine = value

    @property
    def paint_transparent(self) -> bool:
        """
        Gets/Sets whether the control paints it background or not.
        """
        return self.__component.PaintTransparent

    @paint_transparent.setter
    def paint_transparent(self, value: bool) -> None:
        self.__component.PaintTransparent = value

    @property
    def printable(self) -> bool:
        """
        Gets/Sets that the control will be printed with the document.
        """
        return self.__component.Printable

    @printable.setter
    def printable(self, value: bool) -> None:
        self.__component.Printable = value

    @property
    def read_only(self) -> bool:
        """
        Gets/Sets if the content of the control cannot be modified by the user.
        """
        return self.__component.ReadOnly

    @read_only.setter
    def read_only(self, value: bool) -> None:
        self.__component.ReadOnly = value

    @property
    def tabstop(self) -> bool:
        """
        Gets/Sets if the control can be reached with the TAB key.
        """
        return self.__component.Tabstop

    @tabstop.setter
    def tabstop(self, value: bool) -> None:
        self.__component.Tabstop = value

    @property
    def text(self) -> str:
        """
        Gets/Sets the text displayed in the control.
        """
        return self.__component.Text

    @text.setter
    def text(self, value: str) -> None:
        self.__component.Text = value

    @property
    def text_color(self) -> Color:
        """
        Gets/Sets the text color of the control.
        """
        return Color(self.__component.TextColor)

    @text_color.setter
    def text_color(self, value: Color) -> None:
        self.__component.TextColor = value  # type: ignore

    @property
    def text_line_color(self) -> Color:
        """
        Gets/Sets the text line color (RGB) of the control.
        """
        return Color(self.__component.TextLineColor)

    @text_line_color.setter
    def text_line_color(self, value: Color) -> None:
        self.__component.TextLineColor = value  # type: ignore

    @property
    def v_scroll(self) -> bool:
        """
        Gets/Sets if the content of the control can be scrolled in the vertical direction.
        """
        return self.__component.VScroll

    @v_scroll.setter
    def v_scroll(self, value: bool) -> None:
        self.__component.VScroll = value

    @property
    def vertical_align(self) -> VerticalAlignment:
        """
        specifies the vertical alignment of the text in the control.

        Hint:
            - ``VerticalAlignment`` can be imported from ``ooo.dyn.style.vertical_alignment``
        """
        return self.__component.VerticalAlign  # type: ignore

    @vertical_align.setter
    def vertical_align(self, value: VerticalAlignment) -> None:
        self.__component.VerticalAlign = value  # type: ignore

    @property
    def writing_mode(self) -> int:
        """
        Denotes the writing mode used in the control, as specified in the ``com.sun.star.text.WritingMode2`` constants group.

        Only LR_TB (``0``) and RL_TB (``1``) are supported at the moment.
        """
        return self.__component.WritingMode

    @writing_mode.setter
    def writing_mode(self, value: int) -> None:
        self.__component.WritingMode = value

    # endregion Properties
