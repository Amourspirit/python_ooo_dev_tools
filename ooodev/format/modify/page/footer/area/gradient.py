from __future__ import annotations
import uno
from ooo.dyn.awt.gradient_style import GradientStyle as GradientStyle
from ......utils.data_type.angle import Angle as Angle
from ......utils.data_type.color_range import ColorRange as ColorRange
from ......utils.data_type.intensity import Intensity as Intensity
from ......utils.data_type.intensity_range import IntensityRange as IntensityRange
from ......utils.data_type.offset import Offset as Offset
from .....writer.style.page.kind.writer_style_page_kind import WriterStylePageKind as WriterStylePageKind
from .....preset.preset_gradient import PresetGradientKind as PresetGradientKind
from .....direct.common.props.area_gradient_props import AreaGradientProps
from ...header.area.gradient import Gradient as HeaderGradient


class Gradient(HeaderGradient):
    """
    Page Footer Gradient Color

    .. versionadded:: 0.9.0
    """

    # region internal methods
    def _get_inner_props(self) -> AreaGradientProps:
        return AreaGradientProps(
            style="FooterFillStyle",
            step_count="FooterFillGradientStepCount",
            name="FooterFillGradientName",
            grad_prop_name="FooterFillGradient",
        )

    # endregion internal methods
