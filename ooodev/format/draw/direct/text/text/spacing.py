from __future__ import annotations
from typing import TYPE_CHECKING
from ooodev.format.inner.direct.draw.shape.text.text.spacing import Spacing as TextSpacing

if TYPE_CHECKING:
    from ooodev.units.unit_obj import UnitT


class Spacing(TextSpacing):
    """
    This class represents the text spacing of an object that supports ``com.sun.star.drawing.TextProperties``.

    .. seealso::

        - :ref:`help_draw_format_direct_shape_text_text_spacing`

    .. versionadded:: 0.17.5
    """

    def __init__(
        self,
        left: float | UnitT | None = None,
        right: float | UnitT | None = None,
        top: float | UnitT | None = None,
        bottom: float | UnitT | None = None,
    ) -> None:
        """
        Constructor.

        Args:
            left (float, UnitT, optional): Left spacing in MM units or ``UnitT``. Defaults to None.
            right (float, UnitT, optional): Right spacing in MM units or ``UnitT``. Defaults to None.
            top (float, UnitT, optional): Top spacing in MM units or ``UnitT``. Defaults to None.
            bottom (float, UnitT, optional): Bottom spacing in MM units or ``UnitT``. Defaults to None.

        Returns:
            None:

        See Also:
            - :ref:`help_draw_format_direct_shape_text_text_spacing`
        """
        super().__init__(left=left, right=right, top=top, bottom=bottom)
