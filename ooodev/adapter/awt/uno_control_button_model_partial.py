from __future__ import annotations
import contextlib
from typing import Any, cast, TYPE_CHECKING
import uno
from ooo.dyn.text.font_emphasis import FontEmphasisEnum
from ooo.dyn.text.font_relief import FontReliefEnum
from ooo.dyn.style.vertical_alignment import VerticalAlignment
from ooo.dyn.awt.image_align import ImageAlignEnum
from ooo.dyn.awt.image_position import ImagePositionEnum
from ooo.dyn.awt.push_button_type import PushButtonType
from ooodev.utils import info as mInfo
from ooodev.events.events import Events
from ooodev.utils.kind.align_kind import AlignKind
from ooodev.utils.kind.button_state_kind import ButtonStateKind
from ooodev.utils.color import Color
from ooodev.adapter.awt.uno_control_model_partial import UnoControlModelPartial
from ooodev.adapter.awt.font_descriptor_struct_comp import FontDescriptorStructComp
from ooodev.utils.kind.enum_helper import EnumHelper

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlButtonModel  # Service
    from com.sun.star.awt import FontDescriptor  # struct
    from com.sun.star.graphic import XGraphic
    from ooodev.events.args.key_val_args import KeyValArgs


class UnoControlButtonModelPartial(UnoControlModelPartial):
    """Partial class for UnoControlButtonModel."""

    def __init__(self, component: UnoControlButtonModel) -> None:
        """
        Constructor

        Args:
            component (Any): Component that implements ``com.sun.star.awt.UnoControlButtonModel`` service.
        """
        # pylint: disable=unused-argument
        self.__component = component

        UnoControlModelPartial.__init__(self, component=component)
        self.__event_provider = Events(self)
        self.__props = {}

        def on_comp_struct_changed(src: Any, event_args: KeyValArgs) -> None:
            prop_name = str(event_args.event_data["prop_name"])
            if hasattr(self.__component, prop_name):
                setattr(self.__component, prop_name, event_args.source.component)

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
            prop = FontDescriptorStructComp(self.__component.FontDescriptor, key, self.__event_provider)
            self.__props[key] = prop
        return cast(FontDescriptorStructComp, prop)

    @font_descriptor.setter
    def font_descriptor(self, value: FontDescriptor | FontDescriptorStructComp) -> None:
        key = "FontDescriptor"
        if mInfo.Info.is_instance(value, FontDescriptorStructComp):
            self.__component.FontDescriptor = value.copy()
        else:
            self.__component.FontDescriptor = cast("FontDescriptor", value)
        if key in self.__props:
            del self.__props[key]

    @property
    def align(self) -> AlignKind | None:
        """
        Get/Sets the horizontal alignment of the text in the control.

        **optional**

        Hint:
            - ``AlignKind`` can be imported from ``ooodev.utils.kind.align_kind``.
        """
        with contextlib.suppress(AttributeError):
            return AlignKind(self.__component.Align)
        return None

    @align.setter
    def align(self, value: AlignKind | int) -> None:
        kind = AlignKind(int(value))
        with contextlib.suppress(AttributeError):
            self.__component.Align = kind.value

    @property
    def background_color(self) -> Color:
        """
        Gets/Set the background color of the control.

        Returns:
            ~ooodev.utils.color.Color: Color
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
    def image_position(self) -> ImagePositionEnum | None:
        """
        Gets/Sets the position of the image, if any, relative to the text, if any

        Valid values of this property are specified with image_position.

        If this property is present, it supersedes the ImageAlign property - setting one of both properties sets the other one to the best possible match.

        **optional**

        Note:
            Value can be set with ``ImagePositionEnum`` or ``int``.

        Hint:
            - ``ImagePositionEnum`` can be imported from ``ooo.dyn.awt.image_position``
        """
        with contextlib.suppress(AttributeError):
            return ImagePositionEnum(self.__component.ImagePosition)
        return None

    @image_position.setter
    def image_position(self, value: int | ImagePositionEnum) -> None:
        with contextlib.suppress(AttributeError):
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
    def multi_line(self) -> bool | None:
        """
        Gets/Sets that the text may be displayed on more than one line.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.MultiLine
        return None

    @multi_line.setter
    def multi_line(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
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
        # self.__component.PushButtonType is an int
        return cast(
            "PushButtonType",
            EnumHelper.get_uno_enum_from_value("com.sun.star.awt.PushButtonType", self.__component.PushButtonType),
        )

        # values = {0: "STANDARD", 1: "OK", 2: "CANCEL", 3: "HELP"}
        # return uno.Enum("com.sun.star.awt.PushButtonType", values[self.__component.PushButtonType])  # type: ignore

    @push_button_type.setter
    def push_button_type(self, value: int | PushButtonType) -> None:
        val = PushButtonType(value)
        self.__component.PushButtonType = val.value  # type: ignore

    @property
    def repeat(self) -> bool | None:
        """
        Gets/Sets whether the mouse should show repeating behavior, i.e.
        repeatedly trigger an action when keeping pressed.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.Repeat
        return None

    @repeat.setter
    def repeat(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.Repeat = value

    @property
    def repeat_delay(self) -> int | None:
        """
        Gets/Sets the mouse repeat delay, in milliseconds.

        When the user presses a mouse in a control area where this triggers an action (such as spinning the value), then usual control implementations allow to repeatedly trigger this action, without the need to release the mouse button and to press it again.
        The delay between two such triggers is specified with this property.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.RepeatDelay
        return None

    @repeat_delay.setter
    def repeat_delay(self, value: int) -> None:
        with contextlib.suppress(AttributeError):
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

        Returns:
            ~ooodev.utils.color.Color: Color
        """
        return Color(self.__component.TextColor)

    @text_color.setter
    def text_color(self, value: Color) -> None:
        self.__component.TextColor = value  # type: ignore

    @property
    def text_line_color(self) -> Color:
        """
        Gets/Sets the text line color of the control.

        Returns:
            ~ooodev.utils.color.Color: Color
        """
        return Color(self.__component.TextLineColor)

    @text_line_color.setter
    def text_line_color(self, value: Color) -> None:
        self.__component.TextLineColor = value  # type: ignore

    @property
    def toggle(self) -> bool | None:
        """
        Gets/Sets whether the button should toggle on a single operation.

        If this property is set to ``True``, a single operation of the button control (pressing space while it is focused, or clicking onto it)
        toggles it between a pressed and a not pressed state.

        The default for this property is ``False``, which means the button behaves like a usual push button.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.Toggle
        return None

    @toggle.setter
    def toggle(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.Toggle = value

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
