from __future__ import annotations
import uno
from ooodev.format.inner.common.props.area_pattern_props import AreaPatternProps
from ooodev.format.inner.kind.format_kind import FormatKind
from ...header.area.pattern import Pattern as HeaderPattern


class Pattern(HeaderPattern):
    """
    Page Footer Pattern
    .. versionadded:: 0.9.0
    """

    # region Internal Methods
    def _get_inner_props(self) -> AreaPatternProps:
        return AreaPatternProps(
            style="FooterFillStyle",
            name="FooterFillBitmapName",
            tile="FooterFillBitmapTile",
            stretch="FooterFillBitmapStretch",
            bitmap="FooterFillBitmap",
        )

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
