from __future__ import annotations
import contextlib
from typing import Any, TYPE_CHECKING
from ooo.dyn.text.font_emphasis import FontEmphasisEnum
from ooo.dyn.text.font_relief import FontReliefEnum
from ooo.dyn.style.vertical_alignment import VerticalAlignment
from ooo.dyn.awt.mouse_wheel_behavior import MouseWheelBehaviorEnum
from ooodev.utils.kind.border_kind import BorderKind
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.utils.color import Color
from .uno_control_model_partial import UnoControlModelPartial
from .font_descriptor_comp import FontDescriptorComp

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlNumericFieldModel  # Service
    from com.sun.star.awt import FontDescriptor  # struct
    from ooodev.events.args.key_val_args import KeyValArgs


class UnoControlNumericFieldModelPartial(UnoControlModelPartial):
    """Partial class for UnoControlNumericFieldModel."""

    def __init__(self, component: UnoControlNumericFieldModel):
        """
        Constructor

        Args:
            component (Any): Component that implements ``com.sun.star.awt.UnoControlNumericFieldModel`` service.
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
    def decimal_accuracy(self) -> int:
        """
        Gets/Sets the decimal accuracy.
        """
        return self.__component.DecimalAccuracy

    @decimal_accuracy.setter
    def decimal_accuracy(self, value: int) -> None:
        self.__component.DecimalAccuracy = value

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
    def hide_inactive_selection(self) -> bool | None:
        """
        Gets/Sets whether the selection in the control should be hidden when the control is not active (focused).

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.HideInactiveSelection
        return None

    @hide_inactive_selection.setter
    def hide_inactive_selection(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.HideInactiveSelection = value

    @property
    def mouse_wheel_behavior(self) -> MouseWheelBehaviorEnum | None:
        """
        Gets/Sets how the mouse wheel can be used to scroll through the control's content.

        Usually, the mouse wheel scroll through the control's entry list.
        Using this property,you can control under which circumstances this is possible.

        **optional**

        Note:
            Value can be set with ``MouseWheelBehaviorEnum`` or ``int``.

        Hint:
            - ``MouseWheelBehaviorEnum`` can be imported from ``ooo.dyn.awt.mouse_wheel_behavior``
        """
        with contextlib.suppress(AttributeError):
            return MouseWheelBehaviorEnum(self.__component.MouseWheelBehavior)
        return None

    @mouse_wheel_behavior.setter
    def mouse_wheel_behavior(self, value: int | MouseWheelBehaviorEnum) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.MouseWheelBehavior = int(value)

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
    def show_thousands_separator(self) -> bool:
        """
        Gets/Sets whether the thousands separator is to be displayed.
        """
        return self.__component.ShowThousandsSeparator

    @show_thousands_separator.setter
    def show_thousands_separator(self, value: bool) -> None:
        self.__component.ShowThousandsSeparator = value

    @property
    def spin(self) -> bool:
        """
        Gets/Sets if the control has a spin button.
        """
        return self.__component.Spin

    @spin.setter
    def spin(self, value: bool) -> None:
        self.__component.Spin = value

    @property
    def strict_format(self) -> bool:
        """
        Gets/Sets if the value is checked during the user input.
        """
        return self.__component.StrictFormat

    @strict_format.setter
    def strict_format(self, value: bool) -> None:
        self.__component.StrictFormat = value

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
    def value(self) -> float:
        """
        Gets/Sets the value displayed in the control.
        """
        return self.__component.Value

    @value.setter
    def value(self, value: float) -> None:
        self.__component.Value = value

    @property
    def value_max(self) -> float:
        """
        Gets/Sets the maximum value that can be entered.
        """
        return self.__component.ValueMax

    @value_max.setter
    def value_max(self, value: float) -> None:
        self.__component.ValueMax = value

    @property
    def value_min(self) -> float:
        """
        Gets/Sets the minimum value that can be entered.
        """
        return self.__component.ValueMin

    @value_min.setter
    def value_min(self, value: float) -> None:
        self.__component.ValueMin = value

    @property
    def value_step(self) -> float:
        """
        Gets/Sets the value step when using the spin button.
        """
        return self.__component.ValueStep

    @value_step.setter
    def value_step(self, value: float) -> None:
        self.__component.ValueStep = value

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

    @property
    def writing_mode(self) -> int | None:
        """
        Denotes the writing mode used in the control, as specified in the ``com.sun.star.text.WritingMode2`` constants group.

        Only LR_TB (``0``) and RL_TB (``1``) are supported at the moment.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.WritingMode
        return None

    @writing_mode.setter
    def writing_mode(self, value: int) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.WritingMode = value

    # endregion Properties