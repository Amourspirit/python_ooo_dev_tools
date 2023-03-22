from __future__ import annotations
from ooodev.format.inner.common.props.transparent_gradient_props import TransparentGradientProps
from ...header.transparency.gradient import Gradient as HeaderGradient


class Gradient(HeaderGradient):
    """
    Page Footer Transparent Gradient

    .. versionadded:: 0.9.0
    """

    # region Internal Methods
    def _get_inner_props(self) -> TransparentGradientProps:
        return TransparentGradientProps(
            transparence="FooterFillTransparence",
            name="FooterFillTransparenceGradientName",
            struct_prop="FooterFillTransparenceGradient",
        )

    # endregion Internal Methods
