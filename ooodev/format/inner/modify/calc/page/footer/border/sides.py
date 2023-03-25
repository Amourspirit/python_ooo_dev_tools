# region Import
from __future__ import annotations

from ooodev.format.calc.style.page.kind import CalcStylePageKind as CalcStylePageKind
from ooodev.format.inner.direct.structs.side import Side as Side
from ooodev.format.inner.common.props.border_props import BorderProps
from ...header.border.sides import Sides as HeaderSides

# endregion Import


class Sides(HeaderSides):
    """
    Page Footer Style Border Sides.

    .. versionadded:: 0.9.0
    """

    # region overrides
    def _get_inner_props(self) -> BorderProps:
        return BorderProps(
            left="FooterLeftBorder", top="FooterTopBorder", right="FooterRightBorder", bottom="FooterBottomBorder"
        )

    # endregion overrides
