from __future__ import annotations

from ooodev.format.inner.common.props.area_hatch_props import AreaHatchProps
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
