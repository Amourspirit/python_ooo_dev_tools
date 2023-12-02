from __future__ import annotations
from typing import TYPE_CHECKING

from ooodev.format.inner.direct.draw.fill.transparent.transparency import Transparency as FillTransparency

if TYPE_CHECKING:
    from ooodev.utils.data_type.intensity import Intensity


class Transparency(FillTransparency):
    """
    Transparency Transparency

    .. seealso::

        - :ref:`help_draw_format_direct_transparency_transparency`

    .. versionadded:: 0.9.0
    """

    def __init__(self, value: Intensity | int = 0) -> None:
        """
        Constructor

        Args:
            value (Intensity, int, optional): Specifies the transparency value from ``0`` to ``100``.

        Returns:
            None:

        See Also:

            - :ref:`help_draw_format_direct_transparency_transparency`
        """
        super().__init__(value=value)
