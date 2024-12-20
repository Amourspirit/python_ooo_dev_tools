from __future__ import annotations
import contextlib
from typing import TYPE_CHECKING
from ooodev.utils.color import Color
from ooodev.utils.kind.border_kind import BorderKind
from ooodev.adapter.awt.uno_control_model_partial import UnoControlModelPartial

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlProgressBarModel  # Service


class UnoControlProgressBarModelPartial(UnoControlModelPartial):
    """Partial class for UnoControlProgressBarModel."""

    def __init__(self, component: UnoControlProgressBarModel):
        """
        Constructor

        Args:
            component (Any): Component that implements ``com.sun.star.awt.UnoControlProgressBarModel`` service.
        """
        self.__component = component
        # pylint: disable=unused-argument
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
            ~ooodev.utils.color.Color | None: Color or None if not set.
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
    def fill_color(self) -> Color:
        """
        Gets/Set the fill color of the control.
        """
        return Color(self.__component.FillColor)

    @fill_color.setter
    def fill_color(self, value: Color) -> None:
        self.__component.FillColor = value  # type: ignore

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
    def printable(self) -> bool:
        """
        Gets/Sets that the control will be printed with the document.
        """
        return self.__component.Printable

    @printable.setter
    def printable(self, value: bool) -> None:
        self.__component.Printable = value

    @property
    def progress_value(self) -> int:
        """
        Gets/Sets the progress value of the control.
        """
        return self.__component.ProgressValue

    @progress_value.setter
    def progress_value(self, value: int) -> None:
        self.__component.ProgressValue = value

    @property
    def progress_value_max(self) -> int:
        """
        Gets/Sets the maximum progress value of the control.
        """
        return self.__component.ProgressValueMax

    @progress_value_max.setter
    def progress_value_max(self, value: int) -> None:
        self.__component.ProgressValueMax = value

    @property
    def progress_value_min(self) -> int:
        """
        Gets/Sets the minimum progress value of the control.
        """
        return self.__component.ProgressValueMin

    @progress_value_min.setter
    def progress_value_min(self, value: int) -> None:
        self.__component.ProgressValueMin = value

    # endregion Properties
