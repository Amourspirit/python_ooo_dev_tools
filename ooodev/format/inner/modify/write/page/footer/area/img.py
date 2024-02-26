from __future__ import annotations
import uno
from com.sun.star.awt import XBitmap
from ooo.dyn.drawing.rectangle_point import RectanglePoint

from ooodev.format.inner.common.format_types.offset_column import OffsetColumn
from ooodev.format.inner.common.format_types.offset_row import OffsetRow
from ooodev.format.inner.common.format_types.size_percent import SizePercent
from ooodev.format.inner.common.props.area_img_props import AreaImgProps
from ooodev.format.inner.direct.write.fill.area.img import ImgStyleKind
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.modify.write.page.header.area.img import Img as HeaderImg
from ooodev.format.writer.style.page.kind.writer_style_page_kind import WriterStylePageKind
from ooodev.utils.data_type.offset import Offset
from ooodev.utils.data_type.size_mm import SizeMM


class Img(HeaderImg):
    """
    Page Footer Background Image

    .. seealso::

        - :ref:`help_writer_format_modify_page_footer_area`

    .. versionadded:: 0.9.0
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
        style_name: WriterStylePageKind | str = WriterStylePageKind.STANDARD,
        style_family: str = "PageStyles",
    ) -> None:
        """
        Constructor

        Args:
            bitmap (XBitmap, optional): Bitmap instance. If ``name`` is not already in the Bitmap Table then
                this property is required.
            name (str, optional): Specifies the name of the image. This is also the name that is used to store bitmap
                in LibreOffice Bitmap Table.
            mode (ImgStyleKind, optional): Specifies the image style, tiled, stretched etc.
                Default ``ImgStyleKind.TILED``.
            size (SizePercent, SizeMM, optional): Size in percent (``0 - 100``) or size in ``mm`` units.
            position (RectanglePoint): Tiling position of Image.
            pos_offset (Offset, optional): Tiling position offset.
            tile_offset (OffsetColumn, OffsetRow, optional): The tiling offset.
            auto_name (bool, optional): Specifies if ``name`` is ensured to be unique. Defaults to ``False``.
            style_name (StyleParaKind, str, optional): Specifies the Page Style that instance applies to.
                Default is Default Page Style.
            style_family (str, optional): Style family. Default ``PageStyles``.

        Returns:
            None:

        See Also:
            - :ref:`help_writer_format_modify_page_footer_area`
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

    # region internal methods
    def _get_inner_props(self) -> AreaImgProps:
        return AreaImgProps(
            name="FooterFillBitmapName",
            style="FooterFillStyle",
            mode="FooterFillBitmapMode",
            point="FooterFillBitmapRectanglePoint",
            bitmap="FooterFillBitmap",
            offset_x="FooterFillBitmapOffsetX",
            offset_y="FooterFillBitmapOffsetY",
            pos_x="FooterFillBitmapPositionOffsetX",
            pos_y="FooterFillBitmapPositionOffsetY",
            size_x="FooterFillBitmapSizeX",
            size_y="FooterFillBitmapSizeY",
        )

    # endregion internal methods

    # region properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.DOC | FormatKind.STYLE | FormatKind.FOOTER
        return self._format_kind_prop

    # endregion properties
