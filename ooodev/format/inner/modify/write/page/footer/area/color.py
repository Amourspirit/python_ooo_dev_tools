from __future__ import annotations
from ooodev.format.inner.common.props.fill_color_props import FillColorProps
from ...header.area.color import Color as HeaderColor


class Color(HeaderColor):
    """
    Page Footer Color

    .. versionadded:: 0.9.0
    """

    # region internal methods
    def _get_inner_props(self) -> FillColorProps:
        return FillColorProps(color="FooterFillColor", style="FooterFillStyle")

    # endregion internal methods
