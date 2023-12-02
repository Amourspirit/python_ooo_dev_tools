from __future__ import annotations
from typing import TYPE_CHECKING

import uno
from ooodev.format.inner.direct.write.fill.area.pattern import Pattern as FillPattern

if TYPE_CHECKING:
    from com.sun.star.awt import XBitmap


class Pattern(FillPattern):
    """
    Class for Area Fill Pattern.

    .. seealso::

        - :ref:`help_draw_format_direct_shape_area_pattern`

    .. versionadded:: 0.9.3
    """

    def __init__(
        self,
        *,
        bitmap: XBitmap | None = None,
        name: str = "",
        tile: bool = True,
        stretch: bool = False,
        auto_name: bool = False,
    ) -> None:
        """
        Constructor

        Args:
            bitmap (XBitmap, optional): Bitmap instance. If ``name`` is not already in the Bitmap Table then this property is required.
            name (str, optional): Specifies the name of the pattern. This is also the name that is used to store bitmap in LibreOffice Bitmap Table.
            tile (bool, optional): Specified if bitmap is tiled. Defaults to ``True``.
            stretch (bool, optional): Specifies if bitmap is stretched. Defaults to ``False``.
            auto_name (bool, optional): Specifies if ``name`` is ensured to be unique. Defaults to ``False``.

        Returns:
            None:

        Note:
            If ``auto_name`` is ``False`` then a bitmap for a given name is only required the first call.
            All subsequent call of the same name will retrieve the bitmap form the LibreOffice Bitmap Table.

        See Also:

            - :ref:`help_draw_format_direct_shape_area_pattern`
        """
        super().__init__(bitmap=bitmap, name=name, tile=tile, stretch=stretch, auto_name=auto_name)
