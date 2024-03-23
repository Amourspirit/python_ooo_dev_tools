from __future__ import annotations
import contextlib
from typing import TYPE_CHECKING
import uno
from ooo.dyn.awt.image_scale_mode import ImageScaleModeEnum
from ooodev.utils.kind.border_kind import BorderKind
from ooodev.utils.color import Color
from ooodev.adapter.awt.uno_control_model_partial import UnoControlModelPartial

if TYPE_CHECKING:
    from com.sun.star.graphic import XGraphic
    from com.sun.star.awt import UnoControlImageControlModel  # Service


class UnoControlImageControlModelPartial(UnoControlModelPartial):
    """Partial class for UnoControlImageControlModel."""

    def __init__(self, component: UnoControlImageControlModel):
        """
        Constructor

        Args:
            component (Any): Component that implements ``com.sun.star.awt.UnoControlImageControlModel`` service.
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
    def graphic(self) -> XGraphic | None:
        """
        specifies a graphic to be displayed at the button

        If this property is present, it interacts with the ``image_url`` in the following way:

        - If ``image_url`` is set, ``graphic`` will be reset to an object as loaded from the given image URL, or None if ``image_url`` does not point to a valid image file.
        - If ``graphic`` is set, ``image_url`` will be reset to an empty string.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.Graphic
        return None

    @graphic.setter
    def graphic(self, value: XGraphic) -> None:
        with contextlib.suppress(AttributeError):
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
    def image_url(self) -> str:
        """
        Gets/Sets a URL to an image to use for the button.
        """
        return self.__component.ImageURL

    @image_url.setter
    def image_url(self, value: str) -> None:
        self.__component.ImageURL = value

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
    def scale_image(self) -> bool:
        """
        Gets/Sets if the image is automatically scaled to the size of the control.
        """
        return self.__component.ScaleImage

    @scale_image.setter
    def scale_image(self, value: bool) -> None:
        self.__component.ScaleImage = value

    @property
    def scale_mode(self) -> ImageScaleModeEnum | None:
        """
        defines how to scale the image

        If this property is present, it supersedes the ScaleImage property.

        The value of this property is one of the ImageScaleMode constants.

        **optional**

        Note:
            Value can be set with ``ImageScaleModeEnum`` or ``int``.

        Hint:
            - ``ImageScaleModeEnum`` can be imported from ``ooo.dyn.awt.image_scale_mode``
        """
        with contextlib.suppress(AttributeError):
            return ImageScaleModeEnum(self.__component.ScaleMode)
        return None

    @scale_mode.setter
    def scale_mode(self, value: int | ImageScaleModeEnum) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.ScaleMode = int(value)

    @property
    def tabstop(self) -> bool | None:
        """
        Gets/Sets that the control can be reached with the TAB key.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.Tabstop
        return None

    @tabstop.setter
    def tabstop(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.Tabstop = value

    # endregion Properties
