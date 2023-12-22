from __future__ import annotations
import uno

from ooodev.format.inner.direct.draw.shape.text.animation.no_effect import NoEffect as ShapeNoEffect


class NoEffect(ShapeNoEffect):
    """
    This class represents the text Animation of an object that supports ``com.sun.star.drawing.TextProperties``.

    .. seealso::

        - :ref:`help_draw_format_direct_shape_text_animation_animations`

    .. versionadded:: 0.17.6
    """

    def __init__(self) -> None:
        """
        Constructor.

        Returns:
            None:

        See Also:
            - :ref:`help_draw_format_direct_shape_text_animation_animations`
        """
        super().__init__()
