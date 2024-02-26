"""
Module for working with text columns of a UNO object that supports ``com.sun.star.drawing.TextProperties`` service.

.. versionadded:: 0.17.8
"""

from __future__ import annotations
from typing import TYPE_CHECKING

from ooodev.format.inner.direct.draw.shape.text.text_columns import TextColumns as ShapeTextColumns

if TYPE_CHECKING:
    from ooodev.units.unit_obj import UnitT


class TextColumns(ShapeTextColumns):
    """
    This class represents the text columns of a UNO object that supports ``com.sun.star.drawing.TextProperties`` service.

    .. seealso::

        - :ref:`help_draw_format_direct_shape_text_text_columns`
    """

    def __init__(self, col_count: int | None = None, spacing: float | UnitT | None = None) -> None:
        """
        Constructor.

        Args:
            count (int, optional): Number of columns. Defaults to None.
            spacing (float, UnitT, optional): Spacing between columns in MM units or ``UnitT``. Defaults to None.

        Returns:
            None:

        See Also:
            - :ref:`help_draw_format_direct_shape_text_text_columns`
        """
        super().__init__(col_count=col_count, spacing=spacing)
