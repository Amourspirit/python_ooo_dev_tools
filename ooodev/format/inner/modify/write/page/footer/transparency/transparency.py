from __future__ import annotations
import uno
from ooodev.format.inner.common.props.transparent_transparency_props import TransparentTransparencyProps
from ooodev.format.inner.kind.format_kind import FormatKind
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

    # region properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.DOC | FormatKind.STYLE | FormatKind.FOOTER
        return self._format_kind_prop

    # endregion properties
