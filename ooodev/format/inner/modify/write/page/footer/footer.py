from __future__ import annotations
from ooodev.format.inner.common.props.hf_props import HfProps
from ..header.header import Header


class Footer(Header):
    """
    Page Footer Settings

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
