from __future__ import annotations

from ooodev.format.inner.direct.draw.shape.position_size.adapt import Adapt as ShapeAdapt


class Adapt(ShapeAdapt):
    """
    Shape Adapt (only for Shape Text Boxes).

    .. seealso::

        - :ref:`help_draw_format_direct_shape_position_size_position_size_adapt`

    .. versionadded:: 0.17.3
    """

    def __init__(
        self,
        fit_height: bool | None = None,
        fit_width: bool | None = None,
    ) -> None:
        """
        Constructor

        Args:
            fit_height (bool, optional): Expands the width of the object to the width of the text, if the object is smaller than the text.
            fit_width (bool, optional): Expands the height of the object to the height of the text, if the object is smaller than the text.

        Returns:
            None:

        Note:
            Adapt is only available for Text Boxes.

        See Also:
            - :ref:`help_draw_format_direct_shape_position_size_position_size_adapt`

        .. versionadded:: 0.17.3
        """
        super().__init__(fit_height=fit_height, fit_width=fit_width)
