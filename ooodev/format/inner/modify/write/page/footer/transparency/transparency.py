from __future__ import annotations
from ooodev.format.inner.common.props.transparent_transparency_props import TransparentTransparencyProps
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
