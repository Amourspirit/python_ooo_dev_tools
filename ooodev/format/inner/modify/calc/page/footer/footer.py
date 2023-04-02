from __future__ import annotations
from ooodev.format.inner.common.props.hf_props import HfProps
from ooodev.format.inner.kind.format_kind import FormatKind
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
            shared_first="FirstPageFooterIsShared",
            margin_left="FooterLeftMargin",
            margin_right="FooterRightMargin",
            spacing="FooterBodyDistance",
            spacing_dyn="FooterDynamic",
            height="FooterHeight",
            height_auto="FooterIsDynamicHeight",
        )

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.DOC | FormatKind.STYLE | FormatKind.FOOTER
        return self._format_kind_prop

    # endregion Internal Methods
