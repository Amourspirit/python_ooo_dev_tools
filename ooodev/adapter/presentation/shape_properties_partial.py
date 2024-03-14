from __future__ import annotations
from typing import TYPE_CHECKING
import uno


if TYPE_CHECKING:
    from com.sun.star.presentation import Shape
    from ooo.dyn.presentation.animation_effect import AnimationEffect
    from ooo.dyn.presentation.click_action import ClickAction
    from ooo.dyn.presentation.animation_speed import AnimationSpeed
    from ooodev.utils.color import Color


class ShapePropertiesPartial:
    """
    Partial class for Presentation Shape.

    See Also:
        `API Shape <https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1presentation_1_1Shape.html>`_
    """

    def __init__(self, component: Shape) -> None:
        """
        Constructor

        Args:
            component (Shape): UNO Component that implements ``com.sun.star.presentation.Shape`` service.
        """
        self.__component = component

    # region Shape
    @property
    def bookmark(self) -> str:
        """
        Gets/Sets a generic URL for the property ``on_click``.
        """
        return self.__component.Bookmark

    @bookmark.setter
    def bookmark(self, value: str) -> None:
        self.__component.Bookmark = value

    @property
    def dim_color(self) -> Color:
        """
        This is the color for dimming this shape.

        This color is used if the property ``dim_prev`` is ``True`` and ``dim_hide`` is ``False``.

        Returns:
            ~ooodev.utils.color.Color: Color
        """
        return self.__component.DimColor  # type: ignore

    @dim_color.setter
    def dim_color(self, value: Color) -> None:
        self.__component.DimColor = value  # type: ignore

    @property
    def dim_hide(self) -> bool:
        """
        Gets/Sets - If this property and the property ``dim_prev`` are both ``True``, the shape is hidden instead of dimmed to a color.
        """
        return self.__component.DimHide

    @dim_hide.setter
    def dim_hide(self, value: bool) -> None:
        self.__component.DimHide = value

    @property
    def dim_previous(self) -> bool:
        """
        Gets/Sets - If this property is ``True``, this shape is dimmed to the color of property ``dim_color`` after executing its animation effect.
        """
        return self.__component.DimPrevious

    @dim_previous.setter
    def dim_previous(self, value: bool) -> None:
        self.__component.DimPrevious = value

    @property
    def effect(self) -> AnimationEffect:
        """
        Gets/Sets the animation effect of this shape.

        Returns:
            AnimationEffect: Animation effect.

        Hint:
            - ``AnimationEffect`` can be imported from ``ooo.dyn.presentation.animation_effect``.
        """
        return self.__component.Effect  # type: ignore

    @effect.setter
    def effect(self, value: AnimationEffect) -> None:
        self.__component.Effect = value  # type: ignore

    @property
    def is_empty_presentation_object(self) -> bool:
        """
        Gets - If this is a default presentation object and if it is empty, this property is ``True``.
        """
        return self.__component.IsEmptyPresentationObject

    @property
    def is_presentation_object(self) -> bool:
        """
        Gets if this is a presentation object, this property is TRUE.

        Presentation objects are objects like ``TitleTextShape`` and ``OutlinerShape``.
        """
        return self.__component.IsPresentationObject

    @property
    def on_click(self) -> ClickAction:
        """
        Gets/Sets an action performed after the user clicks on this shape.

        Returns:
            ClickAction: Click action.

        Hint:
            - ``ClickAction`` can be imported from ``ooo.dyn.presentation.click_action``.
        """
        return self.__component.OnClick  # type: ignore

    @on_click.setter
    def on_click(self, value: ClickAction) -> None:
        self.__component.OnClick = value  # type: ignore

    @property
    def play_full(self) -> bool:
        """
        Gets/Sets - If this property is ``True``, the sound of this shape is played in full.

        The default behavior is to stop the sound after completing the animation effect.
        """
        return self.__component.PlayFull

    @play_full.setter
    def play_full(self, value: bool) -> None:
        self.__component.PlayFull = value

    @property
    def presentation_order(self) -> int:
        """
        This is the position of this shape in the order of the shapes which can be animated on its page.

        The animations are executed in this order, starting at the shape with the PresentationOrder ``1``.
        You can change the order by changing this number. Setting it to ``1`` makes this shape the first shape in the execution order for the animation effects.
        """
        return self.__component.PresentationOrder

    @presentation_order.setter
    def presentation_order(self, value: int) -> None:
        self.__component.PresentationOrder = value

    @property
    def sound(self) -> str:
        """
        Gets/Sets the URL to a sound file that is played while the animation effect of this shape is running.
        """
        return self.__component.Sound

    @sound.setter
    def sound(self, value: str) -> None:
        self.__component.Sound = value

    @property
    def sound_on(self) -> bool:
        """
        Gets/sets - If this property is set to ``True``, a sound is played while the animation effect is executed.
        """
        return self.__component.SoundOn

    @sound_on.setter
    def sound_on(self, value: bool) -> None:
        self.__component.SoundOn = value

    @property
    def speed(self) -> AnimationSpeed:
        """
        This is the speed of the animation effect.
        """
        return self.__component.Speed  # type: ignore

    @speed.setter
    def speed(self, value: AnimationSpeed) -> None:
        self.__component.Speed = value  # type: ignore

    @property
    def text_effect(self) -> AnimationEffect:
        """
        Gets/Sets the animation effect for the text inside this shape.
        """
        return self.__component.TextEffect  # type: ignore

    @text_effect.setter
    def text_effect(self, value: AnimationEffect) -> None:
        self.__component.TextEffect = value  # type: ignore

    @property
    def verb(self) -> int:
        """
        Gets/Sets an ``OLE2`` verb for the ClickAction VERB in the property ``on_click``.
        """
        return self.__component.Verb

    @verb.setter
    def verb(self, value: int) -> None:
        self.__component.Verb = value

    # endregion Shape
