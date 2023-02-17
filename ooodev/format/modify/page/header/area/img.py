from __future__ import annotations
from typing import Tuple, cast
import uno
from com.sun.star.awt import XBitmap
from ooo.dyn.drawing.rectangle_point import RectanglePoint as RectanglePoint

from ......utils.data_type.angle import Angle as Angle
from ......utils.data_type.color_range import ColorRange as ColorRange
from ......utils.data_type.intensity import Intensity as Intensity
from ......utils.data_type.intensity_range import IntensityRange as IntensityRange
from ......utils.data_type.offset import Offset as Offset
from .....direct.common.format_types.offset_column import OffsetColumn as OffsetColumn
from .....direct.common.format_types.offset_row import OffsetRow as OffsetRow
from .....direct.common.format_types.size_mm import SizeMM as SizeMM
from .....direct.common.format_types.size_percent import SizePercent as SizePercent
from .....direct.common.props.area_img_props import AreaImgProps
from .....direct.fill.area.img import Img as FillImg, ImgStyleKind as ImgStyleKind
from .....kind.format_kind import FormatKind
from .....preset.preset_image import PresetImageKind as PresetImageKind
from .....writer.style.page.kind.style_page_kind import StylePageKind as StylePageKind
from ...page_style_base_multi import PageStyleBaseMulti


class HeaderImg(FillImg):
    """
    Header Image Background

    .. versionadded:: 0.9.0
    """

    def _supported_services(self) -> Tuple[str, ...]:
        return ("com.sun.star.style.PageProperties", "com.sun.star.style.PageStyle")

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        return FormatKind.DOC | FormatKind.STYLE

    @property
    def _props(self) -> AreaImgProps:
        try:
            return self._props_area_img
        except AttributeError:
            self._props_area_img = AreaImgProps(
                name="HeaderFillBitmapName",
                style="HeaderFillStyle",
                mode="HeaderFillBitmapMode",
                point="HeaderFillBitmapRectanglePoint",
                bitmap="HeaderFillBitmap",
                offset_x="HeaderFillBitmapOffsetX",
                offset_y="HeaderFillBitmapOffsetY",
                pos_x="HeaderFillBitmapPositionOffsetX",
                pos_y="HeaderFillBitmapPositionOffsetY",
                size_x="HeaderFillBitmapSizeX",
                size_y="HeaderFillBitmapSizeY",
            )
        return self._props_area_img


class Img(PageStyleBaseMulti):
    """
    Page Header Background Image

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
        style_name: StylePageKind | str = StylePageKind.STANDARD,
        style_family: str = "PageStyles",
    ) -> None:
        """
        Constructor

        Args:
            bitmap (XBitmap, optional): Bitmap instance. If ``name`` is not already in the Bitmap Table then this property is requied.
            name (str, optional): Specifies the name of the image. This is also the name that is used to store bitmap in LibreOffice Bitmap Table.
            mode (ImgStyleKind, optional): Specifies the image style, tiled, stretched etc. Default ``ImgStyleKind.TILED``.
            size (SizePercent, SizeMM, optional): Size in percent (``0 - 100``) or size in ``mm`` units.
            positin (RectanglePoint): Tiling position of Image.
            pos_offset (Offset, optional): Tiling position offset.
            tile_offset (OffsetColumn, OffsetRow, optional): The tiling offset.
            auto_name (bool, optional): Specifies if ``name`` is ensured to be unique. Defaults to ``False``.
            style_name (StyleParaKind, str, optional): Specifies the Page Style that instance applies to. Deftult is Default Page Style.
            style_family (str, optional): Style family. Defatult ``PageStyles``.

        Returns:
            None:
        """

        direct = HeaderImg(
            bitmap=bitmap,
            name=name,
            mode=mode,
            size=size,
            position=position,
            pos_offset=pos_offset,
            tile_offset=tile_offset,
            auto_name=auto_name,
        )
        super().__init__()
        self._style_name = str(style_name)
        self._style_family_name = style_family
        self._set_style("direct", direct, *direct.get_attrs())

    @classmethod
    def from_style(
        cls,
        doc: object,
        style_name: StylePageKind | str = StylePageKind.STANDARD,
        style_family: str = "PageStyles",
    ) -> Img:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Documnet Object.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to. Deftult is Default Paragraph Style.
            style_family (str, optional): Style family. Defatult ``PageStyles``.

        Returns:
            Img: ``Img`` instance from document properties.
        """
        inst = super(Img, cls).__new__(cls)
        inst.__init__(style_name=style_name, style_family=style_family)
        direct = HeaderImg.from_obj(inst.get_style_props(doc))
        inst._set_style("direct", direct, *direct.get_attrs())
        return inst

    @classmethod
    def from_preset(
        cls,
        preset: PresetImageKind,
        style_name: StylePageKind | str = StylePageKind.STANDARD,
        style_family: str = "PageStyles",
    ) -> Img:
        """
        Gets instance from preset.

        Args:
            preset (PresetKind): Preset.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to. Deftult is Default Paragraph Style.
            style_family (str, optional): Style family. Defatult ``PageStyles``.

        Returns:
            Img: ``Img`` instance from preset.
        """
        inst = super(Img, cls).__new__(cls)
        inst.__init__(style_name=style_name, style_family=style_family)
        direct = HeaderImg.from_preset(preset=preset)
        inst._set_style("direct", direct, *direct.get_attrs())
        return inst

    @property
    def prop_style_name(self) -> str:
        """Gets/Sets property Style Name"""
        return self._style_name

    @prop_style_name.setter
    def prop_style_name(self, value: str | StylePageKind):
        self._style_name = str(value)

    @property
    def prop_inner(self) -> HeaderImg:
        """Gets Inner Image instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(HeaderImg, self._get_style_inst("direct"))
        return self._direct_inner
