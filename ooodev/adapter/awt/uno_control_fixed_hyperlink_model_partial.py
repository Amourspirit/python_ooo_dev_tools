from __future__ import annotations
import contextlib
from typing import Any, TYPE_CHECKING
import uno  # pylint: disable=unused-import
from ooo.dyn.text.font_emphasis import FontEmphasisEnum
from ooo.dyn.text.font_relief import FontReliefEnum
from ooo.dyn.style.vertical_alignment import VerticalAlignment
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.utils.kind.align_kind import AlignKind
from ooodev.utils.kind.border_kind import BorderKind
from ooodev.utils.color import Color
from .uno_control_model_partial import UnoControlModelPartial
from .font_descriptor_comp import FontDescriptorComp

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlFixedHyperlinkModel  # Service
    from com.sun.star.awt import FontDescriptor  # struct
    from ooodev.events.args.key_val_args import KeyValArgs


class UnoControlFixedHyperlinkModelPartial(UnoControlModelPartial):
    """Partial class for UnoControlFixedHyperlinkModel."""

    def __init__(self, component: UnoControlFixedHyperlinkModel):
        """
        Constructor

        Args:
            component (Any): Component that implements ``com.sun.star.awt.UnoControlFixedHyperlinkModel`` service.
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
    def border_color(self) -> Color | None:
        """
        Gets/Sets the color of the border, if present

        Not every border style (see Border) may support coloring.
        For instance, usually a border with 3D effect will ignore the border_color setting.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return Color(self.__component.BorderColor)
        return None

    @border_color.setter
    def border_color(self, value: Color) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.BorderColor = value

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
    def label(self) -> str:
        """
        Gets/Sets the label of the control.
        """
        return self.__component.Label

    @label.setter
    def label(self, value: str) -> None:
        self.__component.Label = value

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
    def printable(self) -> bool:
        """
        Gets/Sets that the control will be printed with the document.
        """
        return self.__component.Printable

    @printable.setter
    def printable(self, value: bool) -> None:
        self.__component.Printable = value

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
    def url(self) -> str:
        """
        Gets/Sets the URL.
        """
        return self.__component.URL

    @url.setter
    def url(self, value: str) -> None:
        self.__component.URL = value

    @property
    def vertical_align(self) -> VerticalAlignment | None:
        """
        Gets/Sets the vertical alignment of the text in the control.

        **optional**

        Hint:
            - ``VerticalAlignment`` can be imported from ``ooo.dyn.style.vertical_alignment``
        """
        with contextlib.suppress(AttributeError):
            return self.__component.VerticalAlign  # type: ignore
        return None

    @vertical_align.setter
    def vertical_align(self, value: VerticalAlignment) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.VerticalAlign = value  # type: ignore

    # endregion Properties
