from __future__ import annotations

from ooodev.format.inner.common.props.border_props import BorderProps
from ...header.border.sides import Sides as HeaderSides


class Sides(HeaderSides):
    """
    Page Footer Style Border Sides.

    .. versionadded:: 0.9.0
    """

    # region Internal Methods
    def _get_inner_props(self) -> BorderProps:
        return BorderProps(
            left="FooterLeftBorder", top="FooterTopBorder", right="FooterRightBorder", bottom="FooterBottomBorder"
        )

    # endregion Internal Methods
