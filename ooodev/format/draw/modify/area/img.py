from __future__ import annotations

from typing import TYPE_CHECKING
import uno

from ooodev.format.inner.modify.write.fill.area.img import Img as FillImg
from ooodev.utils.data_type.size_mm import SizeMM
from ooodev.format.inner.direct.write.fill.area.img import ImgStyleKind
from ooodev.format.draw.style.kind import DrawStyleFamilyKind
from ooodev.format.draw.style.lookup import FamilyGraphics


if TYPE_CHECKING:
    from com.sun.star.awt import XBitmap
    from ooo.dyn.drawing.rectangle_point import RectanglePoint
    from ooodev.format.inner.direct.write.fill.area.img import SizePercent
    from ooodev.utils.data_type.offset import Offset
    from ooodev.format.inner.direct.write.fill.area.img import OffsetColumn
    from ooodev.format.inner.direct.write.fill.area.img import OffsetRow


class Img(FillImg):
    """
    Shape Style Modify Area Fill Image

    .. seealso::

        - :ref:`help_draw_format_modify_area_image`

    .. versionadded:: 0.17.9
    """

    def __init__(
        self,
        *,
        bitmap: XBitmap | None = None,
        name: str = "",
        mode: ImgStyleKind = ImgStyleKind.TILED,
        size: SizePercent | SizeMM | None = None,
        position: RectanglePoint | None = None,
        pos_offset: Offset | None = None,
        tile_offset: OffsetColumn | OffsetRow | None = None,
        auto_name: bool = False,
        style_name: str = FamilyGraphics.DEFAULT_DRAWING_STYLE,
        style_family: str | DrawStyleFamilyKind = DrawStyleFamilyKind.GRAPHICS,
    ) -> None:
        """
        Constructor

        Args:
            bitmap (XBitmap, optional): Bitmap instance. If ``name`` is not already in the Bitmap Table
                then this property is required.
            name (str, optional): Specifies the name of the image. This is also the name that is used to store
                bitmap in LibreOffice Bitmap Table.
            mode (ImgStyleKind, optional): Specifies the image style, tiled, stretched etc.
                Default ``ImgStyleKind.TILED``.
            size (SizePercent, SizeMM, optional): Size in percent (``0 - 100``) or size in ``mm`` units.
            position (RectanglePoint): Tiling position of Image.
            pos_offset (Offset, optional): Tiling position offset.
            tile_offset (OffsetColumn, OffsetRow, optional): The tiling offset.
            auto_name (bool, optional): Specifies if ``name`` is ensured to be unique. Defaults to ``False``.
            style_name (FamilyGraphics, str, optional): Specifies the Style that instance applies to.
                Default is Default ``standard`` Style.
            style_family (str, DrawStyleFamilyKind, optional): Family Style. Defaults to ``graphics``.

        Returns:
            None:

        See Also:
            - :ref:`help_draw_format_modify_area_image`
        """
        super().__init__(
            bitmap=bitmap,
            name=name,
            mode=mode,
            size=size,
            position=position,
            pos_offset=pos_offset,
            tile_offset=tile_offset,
            auto_name=auto_name,
            style_name=style_name,
            style_family=style_family,
        )
