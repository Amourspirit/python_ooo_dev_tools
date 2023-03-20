from __future__ import annotations
import uno

from ......utils.data_type.angle import Angle as Angle
from ......utils.data_type.color_range import ColorRange as ColorRange
from ......utils.data_type.intensity import Intensity as Intensity
from ......utils.data_type.intensity_range import IntensityRange as IntensityRange
from ......utils.data_type.offset import Offset as Offset
from .....writer.style.page.kind.writer_style_page_kind import WriterStylePageKind as WriterStylePageKind
from .....preset.preset_pattern import PresetPatternKind as PresetPatternKind
from .....direct.common.props.area_pattern_props import AreaPatternProps
from ...header.area.pattern import Pattern as HeaderPattern


class Pattern(HeaderPattern):
    """
    Page Footer Pattern
    .. versionadded:: 0.9.0
    """

    # region Internal Methods
    def _get_inner_props(self) -> AreaPatternProps:
        return AreaPatternProps(
            style="FooterFillStyle",
            name="FooterFillBitmapName",
            tile="FooterFillBitmapTile",
            stretch="FooterFillBitmapStretch",
            bitmap="FooterFillBitmap",
        )

    # endregion Internal Methods
