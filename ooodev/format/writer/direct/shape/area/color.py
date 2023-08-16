"""Writer Direct Shape Area Color Class."""
from ooodev.format.inner.direct.write.fill.area.fill_color import FillColor
from ooodev.utils import color as mColor


class Color(FillColor):
    """
    Class for Area color.

    .. seealso::

        - :ref:`help_writer_format_direct_shape_color`

    .. versionadded:: 0.9.3
    """

    def __init__(
        self, color: mColor.Color = mColor.StandardColor.AUTO_COLOR  # pylint: disable=redefined-outer-name
    ) -> None:
        """
        Constructor

        Args:
            color (:py:data:`~.utils.color.Color`, optional): FillColor Color.

        Returns:
            None:

        See Also:

            - :ref:`help_writer_format_direct_shape_color`
        """
        super().__init__(color=color)
