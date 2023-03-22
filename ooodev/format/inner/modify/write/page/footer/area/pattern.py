from __future__ import annotations

from ooodev.format.inner.common.props.area_pattern_props import AreaPatternProps
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
