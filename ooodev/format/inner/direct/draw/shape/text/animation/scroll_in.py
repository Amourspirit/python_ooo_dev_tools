from __future__ import annotations
from typing import TYPE_CHECKING, Set
import uno
from ooo.dyn.drawing.text_animation_kind import TextAnimationKind
from ooo.dyn.drawing.text_animation_direction import TextAnimationDirection
from ooodev.format.inner.direct.draw.shape.text.animation.blink import Blink

if TYPE_CHECKING:
    from ooodev.units.unit_obj import UnitT


class ScrollIn(Blink):
    """
    This class represents the text Animation of an object that supports ``com.sun.star.drawing.TextProperties``.
    """

    def __init__(
        self,
        start_inside: bool | None = None,
        visible_on_exit: bool | None = None,
        increment: float | UnitT | None = None,
        delay: int | None = None,
        direction: TextAnimationDirection | None = None,
        count: int | None = None,
    ) -> None:
        """
        Constructor.

        Args:
            start_inside: bool, optional): Start inside.
                If this value is ``True``, the text is visible at the start of the animation.
            visible_on_exit (bool, optional): Visible on exit.
                If this value is ``True``, the text is visible at the end of the animation.
            increment (float, UnitT, optional): Increment value.
                A positive float represents MM units, a negative value represents pixels,
                or a ``UnitT`` instance can be passed in.
            delay (int, optional): Delay in milliseconds.
                This is the delay in thousandths of a second (milliseconds) between each of the animation steps.
                A value of 0 means automatic delay.
            direction (TextAnimationDirection, optional): This enumeration defines the direction in which the text moves.
            count (int, optional): This number defines how many times the text animation is repeated.
                If this is set to ``0``, the repeat is endless.
        """
        super().__init__(start_inside=start_inside, visible_on_exit=visible_on_exit, increment=increment, delay=delay)
        self.prop_direction = direction
        self.prop_count = count

    # region Overridden Methods
    def _set_animation(self) -> None:
        self._set("TextAnimationKind", TextAnimationKind.SLIDE)

    def _get_uno_props(self) -> Set[str]:
        props = super()._get_uno_props()
        props.add("TextAnimationDirection")
        props.add("TextAnimationCount")
        return props

    # endregion Overridden Methods

    @property
    def prop_direction(self) -> TextAnimationDirection | None:
        """This enumeration defines the direction in which the text moves."""
        return self._get("TextAnimationDirection")

    @prop_direction.setter
    def prop_direction(self, value: TextAnimationDirection | None) -> None:
        if value is None:
            self._remove("TextAnimationDirection")
            return
        self._set("TextAnimationDirection", value)

    @property
    def prop_count(self) -> int | None:
        """
        This number defines how many times the text animation is repeated.

        If this is set to ``0``, the repeat is endless.
        """
        return self._get("TextAnimationCount")

    @prop_count.setter
    def prop_count(self, value: int | None) -> None:
        if value is None:
            self._remove("TextAnimationCount")
            return
        if value < 0:
            value = 0
        self._set("TextAnimationCount", value)
