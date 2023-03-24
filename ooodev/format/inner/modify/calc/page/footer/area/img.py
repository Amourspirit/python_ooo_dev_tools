# region Imports
from __future__ import annotations
from ooo.dyn.drawing.rectangle_point import RectanglePoint as RectanglePoint

from ooodev.utils.data_type.offset import Offset as Offset
from ooodev.format.calc.style.page.kind import CalcStylePageKind as CalcStylePageKind

from ooodev.utils.data_type.size_mm import SizeMM as SizeMM
from ooodev.format.inner.common.format_types.size_percent import SizePercent as SizePercent
from ooodev.format.inner.common.format_types.offset_row import OffsetRow as OffsetRow
from ooodev.format.inner.common.format_types.offset_column import OffsetColumn as OffsetColumn
from ooodev.format.inner.preset.preset_image import PresetImageKind as PresetImageKind
from ooodev.format.inner.direct.write.fill.area.img import ImgStyleKind as ImgStyleKind
from ooodev.format.inner.common.props.img_para_area_props import ImgParaAreaProps
from ...header.area.img import Img as HeaderImg


# endregion Imports


class Img(HeaderImg):
    """
    Page Style Image

    .. versionadded:: 0.9.0
    """

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

    # endregion internal methods
