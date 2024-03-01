from __future__ import annotations
import contextlib
import datetime
from typing import Any, cast, TYPE_CHECKING
import uno  # pylint: disable=unused-import
from ooo.dyn.text.font_emphasis import FontEmphasisEnum
from ooo.dyn.text.font_relief import FontReliefEnum
from ooo.dyn.style.vertical_alignment import VerticalAlignment
from ooo.dyn.awt.mouse_wheel_behavior import MouseWheelBehaviorEnum
from ooodev.utils import info as mInfo
from ooodev.events.events import Events
from ooodev.utils.date_time_util import DateUtil
from ooodev.utils.kind.border_kind import BorderKind
from ooodev.utils.kind.date_format_kind import DateFormatKind
from ooodev.utils.color import Color
from ooodev.utils.partial.model_prop_partial import ModelPropPartial
from ooodev.adapter.awt.uno_control_model_partial import UnoControlModelPartial
from ooodev.adapter.awt.font_descriptor_struct_comp import FontDescriptorStructComp

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlDateFieldModel  # Service
    from com.sun.star.awt import FontDescriptor  # struct
    from ooodev.events.args.key_val_args import KeyValArgs


class UnoControlDateFieldModelPartial(UnoControlModelPartial):
    """Partial class for UnoControlDateFieldModel."""

    def __init__(self):
        """
        Constructor

        Args:
            component (Any): Component that implements ``com.sun.star.awt.UnoControlDateFieldModel`` service.
        """
        # pylint: disable=unused-argument
        if not isinstance(self, ModelPropPartial):
            raise TypeError("This class must be used as a mixin that implements ModelPropPartial.")

        self.model: UnoControlDateFieldModel
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
    def border_color(self) -> Color | None:
        """
        Gets/Sets the color of the border, if present

        Not every border style (see Border) may support coloring.
        For instance, usually a border with 3D effect will ignore the border_color setting.

        **optional**

        Returns:
            ~ooodev.utils.color.Color | None: Color or None if no border color is set.
        """
        with contextlib.suppress(AttributeError):
            return Color(self.model.BorderColor)
        return None

    @border_color.setter
    def border_color(self, value: Color) -> None:
        with contextlib.suppress(AttributeError):
            self.model.BorderColor = value

    @property
    def date(self) -> datetime.date:
        """Gets/Sets the date"""
        return DateUtil.uno_date_to_date(self.model.Date)

    @date.setter
    def date(self, value: datetime.date) -> None:
        self.model.Date = DateUtil.date_to_uno_date(value)

    @property
    def date_format(self) -> DateFormatKind:
        """
        Gets/Sets the format.

        Note:
            Value can be set with ``DateFormatKind`` or ``int``.

        Hint:
            - ``DateFormatKind`` can be imported from ``ooodev.utils.kind.date_format_kind``
        """
        return DateFormatKind(self.model.DateFormat)

    @date_format.setter
    def date_format(self, value: DateFormatKind) -> None:
        self.model.DateFormat = value.value

    @property
    def date_max(self) -> datetime.date:
        """Gets/Sets the max date"""
        return DateUtil.uno_date_to_date(self.model.DateMax)

    @property
    def date_min(self) -> datetime.date:
        """Gets/Sets the min date"""
        return DateUtil.uno_date_to_date(self.model.DateMin)

    @date_min.setter
    def date_min(self, value: datetime.date) -> None:
        self.model.DateMin = DateUtil.date_to_uno_date(value)

    @date_max.setter
    def date_max(self, value: datetime.date) -> None:
        self.model.DateMax = DateUtil.date_to_uno_date(value)

    @property
    def date_show_century(self) -> bool:
        """
        Gets/Sets if the date century is displayed.
        """
        return self.model.DateShowCentury

    @date_show_century.setter
    def date_show_century(self, value: bool) -> None:
        self.model.DateShowCentury = value

    @property
    def dropdown(self) -> bool:
        """
        Gets/Sets if the control has a dropdown button.
        """
        return self.model.Dropdown

    @dropdown.setter
    def dropdown(self, value: bool) -> None:
        self.model.Dropdown = value

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
            return MouseWheelBehaviorEnum(self.model.MouseWheelBehavior)
        return None

    @mouse_wheel_behavior.setter
    def mouse_wheel_behavior(self, value: int | MouseWheelBehaviorEnum) -> None:
        with contextlib.suppress(AttributeError):
            self.model.MouseWheelBehavior = int(value)

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
    def repeat(self) -> bool | None:
        """
        Gets/Sets whether the mouse should show repeating behavior, i.e.
        repeatedly trigger an action when keeping pressed.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.model.Repeat
        return None

    @repeat.setter
    def repeat(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.model.Repeat = value

    @property
    def repeat_delay(self) -> int | None:
        """
        Gets/Sets the mouse repeat delay, in milliseconds.

        When the user presses a mouse in a control area where this triggers an action (such as spinning the value), then usual control implementations allow to repeatedly trigger this action, without the need to release the mouse button and to press it again.
        The delay between two such triggers is specified with this property.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.model.RepeatDelay
        return None

    @repeat_delay.setter
    def repeat_delay(self, value: int) -> None:
        with contextlib.suppress(AttributeError):
            self.model.RepeatDelay = value

    @property
    def spin(self) -> bool:
        """
        Gets/Sets if the control has a spin button.
        """
        return self.model.Spin

    @spin.setter
    def spin(self, value: bool) -> None:
        self.model.Spin = value

    @property
    def strict_format(self) -> bool:
        """
        Gets/Sets if the value is checked during the user input.
        """
        return self.model.StrictFormat

    @strict_format.setter
    def strict_format(self, value: bool) -> None:
        self.model.StrictFormat = value

    @property
    def tabstop(self) -> bool:
        """
        Gets/Sets that the control can be reached with the TAB key.
        """
        return self.model.Tabstop

    @tabstop.setter
    def tabstop(self, value: bool) -> None:
        self.model.Tabstop = value

    @property
    def text(self) -> str | None:
        """
        Gets/Sets the text displayed in the control.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.model.Text
        return None

    @text.setter
    def text(self, value: str) -> None:
        with contextlib.suppress(AttributeError):
            self.model.Text = value

    @property
    def text_color(self) -> Color:
        """
        Gets/Sets the text color of the control.

        Returns:
            ~ooodev.utils.color.Color: Color
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
            ~ooodev.utils.color.Color: Color
        """
        return Color(self.model.TextLineColor)

    @text_line_color.setter
    def text_line_color(self, value: Color) -> None:
        self.model.TextLineColor = value  # type: ignore

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
