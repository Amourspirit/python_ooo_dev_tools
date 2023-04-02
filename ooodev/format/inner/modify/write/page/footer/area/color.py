from __future__ import annotations
import uno
from ooodev.format.inner.common.props.fill_color_props import FillColorProps
from ooodev.format.inner.kind.format_kind import FormatKind
from ...header.area.color import Color as HeaderColor


class Color(HeaderColor):
    """
    Page Footer Color

    .. versionadded:: 0.9.0
    """

    # region internal methods
    def _get_inner_props(self) -> FillColorProps:
        return FillColorProps(color="FooterFillColor", style="FooterFillStyle")

    # endregion internal methods

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
