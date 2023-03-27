# region Import
from __future__ import annotations

from ooodev.format.calc.style.page.kind import CalcStylePageKind as CalcStylePageKind
from ooodev.format.inner.common.props.border_props import BorderProps
from ...header.border.padding import Padding as HeaderPadding

# endregion Import


class Padding(HeaderPadding):
    """
    Page Style Footer Border Padding.

    .. versionadded:: 0.9.0
    """

    # region overrides
    def _get_inner_props(self) -> BorderProps:
        return BorderProps(
            left="FooterLeftBorderDistance",
            top="FooterTopBorderDistance",
            right="FooterRightBorderDistance",
            bottom="FooterBottomBorderDistance",
        )


# endregion overrides
