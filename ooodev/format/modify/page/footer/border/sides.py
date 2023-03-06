from __future__ import annotations
import uno

from .....writer.style.page.kind.style_page_kind import StylePageKind as StylePageKind
from .....direct.structs.side import Side as Side, LineSize as LineSize
from .....direct.common.props.border_props import BorderProps
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
