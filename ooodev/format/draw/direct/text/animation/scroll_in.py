from __future__ import annotations
from typing import TYPE_CHECKING
import uno
from ooodev.format.inner.direct.draw.shape.text.animation.scroll_in import ScrollIn as TextScrollIn

if TYPE_CHECKING:
    from ooo.dyn.drawing.text_animation_direction import TextAnimationDirection
    from ooodev.units.unit_obj import UnitT


class ScrollIn(TextScrollIn):
    """
    This class represents the text Animation of an object that supports ``com.sun.star.drawing.TextProperties``.

    .. seealso::

        - :ref:`help_draw_format_direct_shape_text_animation_animations`

    .. versionadded:: 0.17.6
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

        Returns:
            None:

        See Also:
            - :ref:`help_draw_format_direct_shape_text_animation_animations`
        """
        super().__init__(
            start_inside=start_inside,
            visible_on_exit=visible_on_exit,
            increment=increment,
            delay=delay,
            direction=direction,
            count=count,
        )
