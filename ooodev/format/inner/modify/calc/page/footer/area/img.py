# region Imports
from __future__ import annotations
import uno
from com.sun.star.awt import XBitmap
from ooo.dyn.drawing.rectangle_point import RectanglePoint

from ooodev.format.calc.style.page.kind.calc_style_page_kind import CalcStylePageKind
from ooodev.format.inner.common.format_types.offset_column import OffsetColumn
from ooodev.format.inner.common.format_types.offset_row import OffsetRow
from ooodev.format.inner.common.format_types.size_percent import SizePercent
from ooodev.format.inner.common.props.img_para_area_props import ImgParaAreaProps
from ooodev.format.inner.direct.write.fill.area.img import ImgStyleKind
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.modify.calc.page.header.area.img import Img as HeaderImg
from ooodev.utils.data_type.offset import Offset
from ooodev.utils.data_type.size_mm import SizeMM


# endregion Imports


class Img(HeaderImg):
    """
    Page Style Image

    .. seealso::

        - :ref:`help_calc_format_modify_page_footer_background`

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
        style_name: CalcStylePageKind | str = CalcStylePageKind.DEFAULT,
        style_family: str = "PageStyles",
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
            style_name (CalcStylePageKind, str, optional): Specifies the Page Style that instance applies to.
                Default is Default Page Style.
            style_family (str, optional): Style family. Default ``PageStyles``.

        Returns:
            None:

        See Also:
            - :ref:`help_calc_format_modify_page_footer_background`
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
    def _get_inner_cattribs(self) -> dict:
        props = ImgParaAreaProps(
            back_color="FooterBackColor",
            back_graphic="FooterBackGraphic",
            graphic_loc="FooterBackGraphicLocation",
            transparent="FooterBackTransparent",
        )
        return {
            "_props_internal_attributes": props,
            "_supported_services_values": self._supported_services(),
            "_format_kind_prop": self.prop_format_kind,
        }

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.DOC | FormatKind.STYLE | FormatKind.FOOTER
        return self._format_kind_prop

    # endregion internal methods
