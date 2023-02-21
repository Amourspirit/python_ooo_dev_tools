from __future__ import annotations
import uno
from ooo.dyn.awt.gradient_style import GradientStyle as GradientStyle
from ......utils.data_type.angle import Angle as Angle
from ......utils.data_type.color_range import ColorRange as ColorRange
from ......utils.data_type.intensity import Intensity as Intensity
from ......utils.data_type.intensity_range import IntensityRange as IntensityRange
from ......utils.data_type.offset import Offset as Offset
from .....writer.style.page.kind.style_page_kind import StylePageKind as StylePageKind
from .....preset.preset_gradient import PresetGradientKind as PresetGradientKind
from .....direct.common.props.transparent_transparency_props import TransparentTransparencyProps
from ...header.transparency.transparency import Transparency as HeaderTransparency


class Transparency(HeaderTransparency):
    """
    Footer Transparent Transparency

    .. versionadded:: 0.9.0
    """

    # region Internal Methods
    def _get_inner_props(self) -> TransparentTransparencyProps:
        return TransparentTransparencyProps(transparence="FooterFillTransparence")

    # endregion Internal Methods
