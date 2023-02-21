from __future__ import annotations
import uno
from ....writer.style.page.kind.style_page_kind import StylePageKind as StylePageKind
from ....direct.common.props.hf_props import HfProps
from ..header.header import Header


class Footer(Header):
    """
    Page Header Settings

    .. versionadded:: 0.9.0
    """

    # region Internal Methods
    def _get_inner_props(self) -> HfProps:
        return HfProps(
            on="FooterIsOn",
            shared="FooterIsShared",
            shared_first="FirstIsShared",
            margin_left="FooterLeftMargin",
            margin_right="FooterRightMargin",
            spacing="FooterBodyDistance",
            spacing_dyn="FooterDynamicSpacing",
            height="FooterHeight",
            height_auto="FooterIsDynamicHeight",
        )

    # endregion Internal Methods
