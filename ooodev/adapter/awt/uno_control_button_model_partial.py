from __future__ import annotations
from typing import Any, TYPE_CHECKING
from ooo.dyn.text.font_emphasis import FontEmphasisEnum
from ooo.dyn.text.font_relief import FontReliefEnum
from ooo.dyn.style.vertical_alignment import VerticalAlignment
from ooo.dyn.awt.image_align import ImageAlignEnum
from ooo.dyn.awt.image_position import ImagePositionEnum
from ooo.dyn.awt.push_button_type import PushButtonType
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.utils.kind.align_kind import AlignKind
from ooodev.utils.kind.button_state_kind import ButtonStateKind
from ooodev.utils.color import Color
from .uno_control_model_partial import UnoControlModelPartial
from .font_descriptor_comp import FontDescriptorComp

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlButtonModel  # Service
    from com.sun.star.awt import FontDescriptor  # struct
    from com.sun.star.graphic import XGraphic
    from ooodev.events.args.key_val_args import KeyValArgs


class UnoControlButtonModelPartial(UnoControlModelPartial):
    """Partial class for UnoControlButtonModel."""

    def __init__(self, component: UnoControlButtonModel):
        """
        Constructor

        Args:
            component (Any): Component that implements ``com.sun.star.awt.UnoControlButtonModel`` service.
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
    def default_button(self) -> bool:
        """
        Gets/Sets that the button is the default button on the document.
        """
        return self.__component.DefaultButton

    @default_button.setter
    def default_button(self, value: bool) -> None:
        self.__component.DefaultButton = value

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
    def focus_on_click(self) -> bool:
        """
        Gets/Sets whether the button control should grab the focus when clicked.

        If set to ``True`` (which is the default), the button control automatically grabs the focus when the user clicks onto it with the mouse.
        If set to ``False``, the focus is preserved when the user operates the button control with the mouse.
        """
        return self.__component.FocusOnClick

    @focus_on_click.setter
    def focus_on_click(self, value: bool) -> None:
        self.__component.FocusOnClick = value

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
    def graphic(self) -> XGraphic:
        """
        specifies a graphic to be displayed at the button

        If this property is present, it interacts with the ``image_url`` in the following way:

        - If ``image_url`` is set, ``graphic`` will be reset to an object as loaded from the given image URL, or None if ``image_url`` does not point to a valid image file.
        - If ``graphic`` is set, ``image_url`` will be reset to an empty string.
        """
        return self.__component.Graphic

    @graphic.setter
    def graphic(self, value: XGraphic) -> None:
        self.__component.Graphic = value

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
    def image_align(self) -> ImageAlignEnum:
        """
        Gets/Sets the alignment of the image inside the button as ``image_align`` value.

        Note:
            Value can be set with ``ImageAlignEnum`` or ``int``.

        Hint:
            - ``ImageAlignEnum`` can be imported from ``ooo.dyn.awt.image_align``
        """
        return ImageAlignEnum(self.__component.ImageAlign)

    @image_align.setter
    def image_align(self, value: int | ImageAlignEnum) -> None:
        self.__component.ImageAlign = int(value)

    @property
    def image_position(self) -> ImagePositionEnum:
        """
        Gets/Sets the position of the image, if any, relative to the text, if any

        Valid values of this property are specified with image_position.

        If this property is present, it supersedes the ImageAlign property - setting one of both properties sets the other one to the best possible match.

        Note:
            Value can be set with ``ImagePositionEnum`` or ``int``.

        Hint:
            - ``ImagePositionEnum`` can be imported from ``ooo.dyn.awt.image_position``
        """
        return ImagePositionEnum(self.__component.ImagePosition)

    @image_position.setter
    def image_position(self, value: int | ImagePositionEnum) -> None:
        self.__component.ImagePosition = int(value)

    @property
    def image_url(self) -> str:
        """
        Gets/Sets a URL to an image to use for the button.
        """
        return self.__component.ImageURL

    @image_url.setter
    def image_url(self, value: str) -> None:
        self.__component.ImageURL = value

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
    def push_button_type(self) -> PushButtonType:
        """
        Gets/Sets the default action of the button as push_button_type value.

        Note:
            Value can be set with ``PushButtonType`` or ``int``.

        Hint:
            - ``PushButtonType`` can be imported from ``ooo.dyn.awt.push_button_type``
        """
        return PushButtonType(self.__component.PushButtonType)

    @push_button_type.setter
    def push_button_type(self, value: int | PushButtonType) -> None:
        val = PushButtonType(value)
        self.__component.PushButtonType = val.value  # type: ignore

    @property
    def repeat(self) -> bool:
        """
        Gets/Sets whether the control should show repeating behavior.

        Normally, when you click a button with the mouse, you need to release the mouse button, and press it again. With this property set to TRUE, the button is repeatedly pressed while you hold down the mouse button.
        """
        return self.__component.Repeat

    @repeat.setter
    def repeat(self, value: bool) -> None:
        self.__component.Repeat = value

    @property
    def repeat_delay(self) -> int:
        """
        Gets/Sets the mouse repeat delay, in milliseconds.

        When the user presses a mouse in a control area where this triggers an action (such as pressing the button),
        then usual control implementations allow to repeatedly trigger this action, without the need to release the mouse button and to press it again.
        The delay between two such triggers is specified with this property.
        """
        return self.__component.RepeatDelay

    @repeat_delay.setter
    def repeat_delay(self, value: int) -> None:
        self.__component.RepeatDelay = value

    @property
    def state(self) -> ButtonStateKind:
        """
        Gets/Sets the state of the control.

        If Toggle property is set to ``True``, the pressed state is enabled and its pressed state can be obtained with this property.

        Note:
            Value can be set with ``ButtonStateKind`` or ``int``.

        Hint:
            - ``ButtonStateKind`` can be imported from ``ooodev.utils.kind.button_state_kind``.
        """
        return ButtonStateKind(self.__component.State)

    @state.setter
    def state(self, value: int | ButtonStateKind) -> None:
        self.__component.State = int(value)

    @property
    def tabstop(self) -> bool:
        """
        Gets/Sets that the control can be reached with the TAB key.
        """
        return self.__component.Tabstop

    @tabstop.setter
    def tabstop(self, value: bool) -> None:
        self.__component.Tabstop = value

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
        Gets/Sets the text line color of the control.
        """
        return Color(self.__component.TextLineColor)

    @text_line_color.setter
    def text_line_color(self, value: Color) -> None:
        self.__component.TextLineColor = value  # type: ignore

    @property
    def toggle(self) -> bool:
        """
        Gets/Sets whether the button should toggle on a single operation.

        If this property is set to ``True``, a single operation of the button control (pressing space while it is focused, or clicking onto it)
        toggles it between a pressed and a not pressed state.

        The default for this property is ``False``, which means the button behaves like a usual push button.

        **since**

            OOo 2.0
        """
        return self.__component.Toggle

    @toggle.setter
    def toggle(self, value: bool) -> None:
        self.__component.Toggle = value

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

    # endregion Properties
