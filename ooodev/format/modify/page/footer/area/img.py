from __future__ import annotations
import uno
from ooo.dyn.drawing.rectangle_point import RectanglePoint as RectanglePoint

from ......utils.data_type.angle import Angle as Angle
from ......utils.data_type.color_range import ColorRange as ColorRange
from ......utils.data_type.intensity import Intensity as Intensity
from ......utils.data_type.intensity_range import IntensityRange as IntensityRange
from ......utils.data_type.offset import Offset as Offset
from .....direct.common.format_types.offset_column import OffsetColumn as OffsetColumn
from .....direct.common.format_types.offset_row import OffsetRow as OffsetRow
from ......utils.data_type.size_mm import SizeMM as SizeMM
from .....direct.common.format_types.size_percent import SizePercent as SizePercent
from .....direct.common.props.area_img_props import AreaImgProps
from .....direct.fill.area.img import ImgStyleKind as ImgStyleKind
from .....preset.preset_image import PresetImageKind as PresetImageKind
from .....writer.style.page.kind.writer_style_page_kind import WriterStylePageKind as WriterStylePageKind
from ...header.area.img import Img as HeaderImg


class Img(HeaderImg):
    """
    Page Footer Background Image

    .. versionadded:: 0.9.0
    """

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
