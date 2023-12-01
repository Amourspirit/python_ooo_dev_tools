"""Internal Draw Area Color"""
# pylint: disable=unnecessary-pass
# pylint: disable=unused-import
from __future__ import annotations
import uno
from ooodev.format.inner.direct.write.fill.area.fill_color import FillColor
from ooodev.utils import color as mColor


class Color(FillColor):
    """
    Class for Area color.

    .. seealso::

        - :ref:`help_draw_format_direct_shape_area_color`

    .. versionadded:: 0.9.3
    """

    def __init__(self, color: mColor.Color = mColor.StandardColor.AUTO_COLOR) -> None:
        """
        Constructor

        Args:
            color (:py:data:`~.utils.color.Color`, optional): FillColor Color.

        Returns:
            None:

        See Also:

            - :ref:`help_draw_format_direct_shape_area_color`
        """
        super().__init__(color=color)
