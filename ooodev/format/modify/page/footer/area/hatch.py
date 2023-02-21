from __future__ import annotations
from typing import Tuple, cast
import uno
from ooo.dyn.drawing.hatch_style import HatchStyle as HatchStyle

from ......utils.data_type.angle import Angle as Angle
from ......utils.data_type.color_range import ColorRange as ColorRange
from ......utils.data_type.intensity import Intensity as Intensity
from ......utils.data_type.intensity_range import IntensityRange as IntensityRange
from ......utils.data_type.offset import Offset as Offset
from .....writer.style.page.kind.style_page_kind import StylePageKind as StylePageKind
from .....preset.preset_hatch import PresetHatchKind as PresetHatchKind
from .....direct.common.props.area_hatch_props import AreaHatchProps
from ...header.area.hatch import Hatch as HeaderHatch


class Hatch(HeaderHatch):
    """
    Page Footer Hatch
    .. versionadded:: 0.9.0
    """

    # region internal methods
    def _get_inner_props(self) -> AreaHatchProps:
        return AreaHatchProps(
            color="FooterFillColor",
            style="FooterFillStyle",
            bg="FooterFillBackground",
            hatch_prop="FooterFillHatch",
        )

    # endregion internal methods
