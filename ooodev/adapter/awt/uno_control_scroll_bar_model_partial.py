from __future__ import annotations
import contextlib
from typing import TYPE_CHECKING
import uno  # pylint: disable=unused-import
from ooodev.utils.color import Color
from ooodev.utils.kind.border_kind import BorderKind
from ooodev.utils.kind.orientation_kind import OrientationKind
from ooodev.utils.partial.model_prop_partial import ModelPropPartial
from ooodev.adapter.awt.uno_control_model_partial import UnoControlModelPartial

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlScrollBarModel  # Service


class UnoControlScrollBarModelPartial(UnoControlModelPartial):
    """Partial class for UnoControlScrollBarModel."""

    def __init__(self):
        """
        Constructor

        Args:
            component (Any): Component that implements ``com.sun.star.awt.UnoControlScrollBarModel`` service.
        """
        # pylint: disable=unused-argument
        if not isinstance(self, ModelPropPartial):
            raise TypeError("This class must be used as a mixin that implements ModelPropPartial.")

        self.model: UnoControlScrollBarModel
        UnoControlModelPartial.__init__(self)

    # region Properties
    @property
    def background_color(self) -> Color | None:
        """
        Gets/Set the background color of the control.

        **optional**

        Returns:
            ~ooodev.utils.color.Color | None: Color or None if not present.
        """
        with contextlib.suppress(AttributeError):
            return Color(self.model.BackgroundColor)
        return None

    @background_color.setter
    def background_color(self, value: Color) -> None:
        with contextlib.suppress(AttributeError):
            self.model.BackgroundColor = value  # type: ignore

    @property
    def block_increment(self) -> int:
        """
        Gets/Sets the increment for a block move.
        """
        return self.model.BlockIncrement

    @block_increment.setter
    def block_increment(self, value: int) -> None:
        self.model.BlockIncrement = value

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
            ~ooodev.utils.color.Color | None: Color or None if not present.
        """
        with contextlib.suppress(AttributeError):
            return Color(self.model.BorderColor)
        return None

    @border_color.setter
    def border_color(self, value: Color) -> None:
        with contextlib.suppress(AttributeError):
            self.model.BorderColor = value

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
    def line_increment(self) -> int:
        """
        Gets/Sets the increment for a single line move.
        """
        return self.model.LineIncrement

    @line_increment.setter
    def line_increment(self, value: int) -> None:
        self.model.LineIncrement = value

    @property
    def live_scroll(self) -> bool | None:
        """
        Gets/Sets the scrolling behavior of the control.

        TRUE means, that when the user moves the slider in the scroll bar, the content of the window is updated immediately. FALSE means, that the window is only updated after the user has released the mouse button.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.model.LiveScroll
        return None

    @live_scroll.setter
    def live_scroll(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.model.LiveScroll = value

    @property
    def orientation(self) -> OrientationKind:
        """
        Gets/Sets the orientation of the control.

        Note:
            Value can be set with ``OrientationKind`` or ``int``.

        Hint:
            - ``OrientationKind`` can be imported from ``ooodev.utils.kind.orientation_kind``.
        """
        return OrientationKind(self.model.Orientation)

    @orientation.setter
    def orientation(self, value: int | OrientationKind) -> None:
        self.model.Orientation = int(value)

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
    def scroll_value(self) -> int:
        """
        Gets/Sets the scroll value of the control.
        """
        return self.model.ScrollValue

    @scroll_value.setter
    def scroll_value(self, value: int) -> None:
        self.model.ScrollValue = value

    @property
    def scroll_value_max(self) -> int:
        """
        Gets/Sets the maximum scroll value of the control.
        """
        return self.model.ScrollValueMax

    @scroll_value_max.setter
    def scroll_value_max(self, value: int) -> None:
        self.model.ScrollValueMax = value

    @property
    def scroll_value_min(self) -> int | None:
        """
        Gets/Sets the minimum scroll value of the control.

        If this optional property is not present, clients of the component should assume a minimal scroll value of ``0``.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.model.ScrollValueMin
        return None

    @scroll_value_min.setter
    def scroll_value_min(self, value: int) -> None:
        with contextlib.suppress(AttributeError):
            self.model.ScrollValueMin = value

    @property
    def tabstop(self) -> bool | None:
        """
        Gets/Sets that the control can be reached with the TAB key.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.model.Tabstop
        return None

    @tabstop.setter
    def tabstop(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.model.Tabstop = value

    @property
    def visible_size(self) -> int | None:
        """
        Gets/Sets the visible size of the scroll bar.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.model.VisibleSize
        return None

    @visible_size.setter
    def visible_size(self, value: int) -> None:
        with contextlib.suppress(AttributeError):
            self.model.VisibleSize = value

    # endregion Properties
