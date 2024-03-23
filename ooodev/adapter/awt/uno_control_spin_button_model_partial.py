from __future__ import annotations
import contextlib
from typing import TYPE_CHECKING
import uno  # pylint: disable=unused-import
from ooo.dyn.awt.mouse_wheel_behavior import MouseWheelBehaviorEnum
from ooodev.utils.color import Color
from ooodev.utils.kind.border_kind import BorderKind
from ooodev.utils.kind.orientation_kind import OrientationKind
from ooodev.adapter.awt.uno_control_model_partial import UnoControlModelPartial

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlSpinButtonModel  # Service


class UnoControlSpinButtonModelPartial(UnoControlModelPartial):
    """Partial class for UnoControlSpinButtonModel."""

    def __init__(self, component: UnoControlSpinButtonModel):
        """
        Constructor

        Args:
            component (Any): Component that implements ``com.sun.star.awt.UnoControlSpinButtonModel`` service.
        """
        # pylint: disable=unused-argument
        self.__component = component
        UnoControlModelPartial.__init__(self, component=component)

    # region Properties
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

        Returns:
            ~ooodev.utils.color.Color | None: Color or None if not present.
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
    def orientation(self) -> OrientationKind:
        """
        Gets/Sets the orientation of the control.

        Note:
            Value can be set with ``OrientationKind`` or ``int``.

        Hint:
            - ``OrientationKind`` can be imported from ``ooodev.utils.kind.orientation_kind``.
        """
        return OrientationKind(self.__component.Orientation)

    @orientation.setter
    def orientation(self, value: int | OrientationKind) -> None:
        self.__component.Orientation = int(value)

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
    def repeat(self) -> bool:
        """
        Gets/Sets whether the mouse should show repeating behavior, i.e.
        repeatedly trigger an action when keeping pressed.
        """
        return self.__component.Repeat

    @repeat.setter
    def repeat(self, value: bool) -> None:
        self.__component.Repeat = value

    @property
    def repeat_delay(self) -> int:
        """
        Gets/Sets the mouse repeat delay, in milliseconds.

        When the user presses a mouse in a control area where this triggers an action (such as spinning the value), then usual control implementations allow to repeatedly trigger this action, without the need to release the mouse button and to press it again.
        The delay between two such triggers is specified with this property.
        """
        return self.__component.RepeatDelay

    @repeat_delay.setter
    def repeat_delay(self, value: int) -> None:
        self.__component.RepeatDelay = value

    @property
    def spin_increment(self) -> int:
        """
        Gets/Sets the increment by which the value is changed when using operating the spin button.
        """
        return self.__component.SpinIncrement

    @spin_increment.setter
    def spin_increment(self, value: int) -> None:
        self.__component.SpinIncrement = value

    @property
    def spin_value(self) -> int:
        """
        Gets/Sets the current value of the control.
        """
        return self.__component.SpinValue

    @spin_value.setter
    def spin_value(self, value: int) -> None:
        self.__component.SpinValue = value

    @property
    def spin_value_max(self) -> int:
        """
        Gets/Sets the maximum value of the control.
        """
        return self.__component.SpinValueMax

    @spin_value_max.setter
    def spin_value_max(self, value: int) -> None:
        self.__component.SpinValueMax = value

    @property
    def spin_value_min(self) -> int:
        """
        Gets/Sets the minimum value of the control.
        """
        return self.__component.SpinValueMin

    @spin_value_min.setter
    def spin_value_min(self, value: int) -> None:
        self.__component.SpinValueMin = value

    @property
    def symbol_color(self) -> Color:
        """
        Gets/Sets the color to be used when painting symbols which are part of the control's appearance, such as the arrow buttons.

        Returns:
            ~ooodev.utils.color.Color: Color
        """
        return Color(self.__component.SymbolColor)

    @symbol_color.setter
    def symbol_color(self, value: Color) -> None:
        self.__component.SymbolColor = value  # type: ignore

    # endregion Properties
