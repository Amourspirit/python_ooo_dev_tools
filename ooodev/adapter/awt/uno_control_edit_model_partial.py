from __future__ import annotations
import contextlib
from typing import Any, cast, TYPE_CHECKING
from ooo.dyn.text.font_emphasis import FontEmphasisEnum
from ooo.dyn.text.font_relief import FontReliefEnum
from ooo.dyn.style.vertical_alignment import VerticalAlignment
from ooo.dyn.awt.line_end_format import LineEndFormatEnum
from ooodev.utils import info as mInfo
from ooodev.events.events import Events
from ooodev.utils.kind.align_kind import AlignKind
from ooodev.utils.kind.border_kind import BorderKind
from ooodev.utils.color import Color
from ooodev.utils.partial.model_prop_partial import ModelPropPartial
from ooodev.adapter.awt.uno_control_model_partial import UnoControlModelPartial
from ooodev.adapter.awt.font_descriptor_struct_comp import FontDescriptorStructComp

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlEditModel  # Service
    from com.sun.star.awt import FontDescriptor  # struct
    from ooodev.events.args.key_val_args import KeyValArgs


class UnoControlEditModelPartial(UnoControlModelPartial):
    """Partial class for UnoControlEditModel."""

    def __init__(self):
        """
        Constructor

        Args:
            component (Any): Component that implements ``com.sun.star.awt.UnoControlEditModel`` service.
        """
        if not isinstance(self, ModelPropPartial):
            raise TypeError("This class must be used as a mixin that provides the ModelPropPartial.")
        # pylint: disable=unused-argument
        self.model: UnoControlEditModel
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
    def align(self) -> AlignKind:
        """
        Get/Sets the horizontal alignment of the text in the control.

        Hint:
            - ``AlignKind`` can be imported from ``ooodev.utils.kind.align_kind``.
        """
        return AlignKind(self.model.Align)

    @align.setter
    def align(self, value: AlignKind | int) -> None:
        kind = AlignKind(int(value))
        self.model.Align = kind.value

    @property
    def auto_h_scroll(self) -> bool | None:
        """
        Gets/Sets - If set to ``True`` an horizontal scrollbar will be added automatically when needed.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.model.AutoHScroll
        return None

    @auto_h_scroll.setter
    def auto_h_scroll(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.model.AutoHScroll = value

    @property
    def auto_v_scroll(self) -> bool | None:
        """
        Gets/Sets - If set to ``True`` a vertical scrollbar will be added automatically when needed.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.model.AutoVScroll
        return None

    @auto_v_scroll.setter
    def auto_v_scroll(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.model.AutoVScroll = value

    @property
    def background_color(self) -> Color:
        """
        Gets/Set the background color of the control.

        Returns:
            ~ooodev.utils.color.Color: Color
        """
        return Color(self.model.BackgroundColor)

    @background_color.setter
    def background_color(self, value: Color) -> None:
        self.model.BackgroundColor = value  # type: ignore

    @property
    def border(self) -> BorderKind:
        """
        Gets/Sets the border style of the control.

        Note:
            Value can be set with ``BorderKind`` or ``int``.

        Hint:
            - ``BorderKind`` can be imported from ``ooodev.utils.kind.border_kind``.
        """
        return BorderKind(self.model.Border)

    @border.setter
    def border(self, value: int | BorderKind) -> None:
        kind = BorderKind(int(value))
        self.model.Border = kind.value

    @property
    def border_color(self) -> Color:
        """
        Gets/Sets the color of the border, if present

        Not every border style (see Border) may support coloring. For instance, usually a border with 3D effect will ignore the border_color setting.

        Returns:
            ~ooodev.utils.color.Color: Color
        """
        return Color(self.model.BorderColor)

    @border_color.setter
    def border_color(self, value: Color) -> None:
        self.model.BorderColor = value

    @property
    def echo_char(self) -> str:
        """Gets/Sets the echo character as a string"""
        with contextlib.suppress(Exception):
            return chr(self.model.EchoChar)
        return ""

    @echo_char.setter
    def echo_char(self, value: str) -> None:
        if value != "":
            value = value[0]  # first char
        self.model.EchoChar = ord(value)
        ...

    @property
    def enabled(self) -> bool:
        """
        Gets/Sets whether the control is enabled or disabled.
        """
        return self.model.Enabled

    @enabled.setter
    def enabled(self, value: bool) -> None:
        self.model.Enabled = value

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
    def h_scroll(self) -> bool:
        """
        Gets/Sets if the content of the control can be scrolled in the horizontal direction.
        """
        return self.model.HScroll

    @h_scroll.setter
    def h_scroll(self, value: bool) -> None:
        self.model.HScroll = value

    @property
    def hard_line_breaks(self) -> bool:
        """
        Gets/Sets if hard line breaks will be returned in the ``XTextComponent.getText()`` method.
        """
        return self.model.HardLineBreaks

    @hard_line_breaks.setter
    def hard_line_breaks(self, value: bool) -> None:
        self.model.HardLineBreaks = value

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
    def hide_inactive_selection(self) -> bool | None:
        """
        Gets/Sets whether the selection in the control should be hidden when the control is not active (focused).

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.model.HideInactiveSelection
        return None

    @hide_inactive_selection.setter
    def hide_inactive_selection(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.model.HideInactiveSelection = value

    @property
    def line_end_format(self) -> LineEndFormatEnum | None:
        """
        specifies which line end type should be used for multi line text

        Controls working with this model care for this setting when the user enters text. Every line break entered into the control will be treated according to this setting, so that the Text property always contains only line ends in the format specified.

        Possible values are all constants from the LineEndFormat group.

        Note that this setting is usually not relevant when you set new text via the API. No matter which line end format is used in this new text then, usual control implementations should recognize all line end formats and display them properly.

        **optional**

        Note:
            Value can be set with ``LineEndFormatEnum`` or ``int``.

        Hint:
            - ``LineEndFormatEnum`` can be imported from ``ooo.dyn.awt.line_end_format``.
        """
        with contextlib.suppress(AttributeError):
            return LineEndFormatEnum(self.model.LineEndFormat)
        return None

    @line_end_format.setter
    def line_end_format(self, value: int | LineEndFormatEnum) -> None:
        with contextlib.suppress(AttributeError):
            self.model.LineEndFormat = int(value)

    @property
    def max_text_len(self) -> int:
        """
        Gets/Sets the maximum character count.

        There's no limitation, if set to 0.
        """
        return self.model.MaxTextLen

    @max_text_len.setter
    def max_text_len(self, value: int) -> None:
        self.model.MaxTextLen = value

    @property
    def multi_line(self) -> bool:
        """
        Gets/Sets that the text may be displayed on more than one line.
        """
        return self.model.MultiLine

    @multi_line.setter
    def multi_line(self, value: bool) -> None:
        self.model.MultiLine = value

    @property
    def paint_transparent(self) -> bool | None:
        """
        Gets/Sets whether the control paints it background or not.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.model.PaintTransparent
        return None

    @paint_transparent.setter
    def paint_transparent(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.model.PaintTransparent = value

    @property
    def printable(self) -> bool:
        """
        Gets/Sets that the control will be printed with the document.
        """
        return self.model.Printable

    @printable.setter
    def printable(self, value: bool) -> None:
        self.model.Printable = value

    @property
    def read_only(self) -> bool:
        """
        Gets/Sets if the content of the control cannot be modified by the user.
        """
        return self.model.ReadOnly

    @read_only.setter
    def read_only(self, value: bool) -> None:
        self.model.ReadOnly = value

    @property
    def tabstop(self) -> bool:
        """
        Gets/Sets if the control can be reached with the TAB key.
        """
        return self.model.Tabstop

    @tabstop.setter
    def tabstop(self, value: bool) -> None:
        self.model.Tabstop = value

    @property
    def text(self) -> str:
        """
        Gets/Sets the text displayed in the control.
        """
        return self.model.Text

    @text.setter
    def text(self, value: str) -> None:
        self.model.Text = value

    @property
    def text_color(self) -> Color:
        """
        Gets/Sets the text color of the control.
        """
        return Color(self.model.TextColor)

    @text_color.setter
    def text_color(self, value: Color) -> None:
        self.model.TextColor = value  # type: ignore

    @property
    def text_line_color(self) -> Color:
        """
        Gets/Sets the text line color (RGB) of the control.
        """
        return Color(self.model.TextLineColor)

    @text_line_color.setter
    def text_line_color(self, value: Color) -> None:
        self.model.TextLineColor = value  # type: ignore

    @property
    def v_scroll(self) -> bool:
        """
        Gets/Sets if the content of the control can be scrolled in the vertical direction.
        """
        return self.model.VScroll

    @v_scroll.setter
    def v_scroll(self, value: bool) -> None:
        self.model.VScroll = value

    @property
    def vertical_align(self) -> VerticalAlignment | None:
        """
        Gets/Sets the vertical alignment of the text in the control.

        **optional**

        Hint:
            - ``VerticalAlignment`` can be imported from ``ooo.dyn.style.vertical_alignment``
        """
        with contextlib.suppress(AttributeError):
            return self.model.VerticalAlign  # type: ignore
        return None

    @vertical_align.setter
    def vertical_align(self, value: VerticalAlignment) -> None:
        with contextlib.suppress(AttributeError):
            self.model.VerticalAlign = value  # type: ignore

    @property
    def writing_mode(self) -> int | None:
        """
        Denotes the writing mode used in the control, as specified in the ``com.sun.star.text.WritingMode2`` constants group.

        Only LR_TB (``0``) and RL_TB (``1``) are supported at the moment.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.model.WritingMode
        return None

    @writing_mode.setter
    def writing_mode(self, value: int) -> None:
        with contextlib.suppress(AttributeError):
            self.model.WritingMode = value

    # endregion Properties
