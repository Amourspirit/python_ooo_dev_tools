from __future__ import annotations
from ooodev.format.inner.common.props.area_gradient_props import AreaGradientProps
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
