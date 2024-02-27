# region Import
from __future__ import annotations
import uno
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.common.props.fill_color_props import FillColorProps
from ooodev.format.inner.direct.write.page.header.area.color import Color as HeaderColor

# endregion Import


class Color(HeaderColor):
    """
    Color of the footer area.

    .. versionadded:: 0.9.2
    """

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.FOOTER
        return self._format_kind_prop

    @property
    def _props(self) -> FillColorProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = FillColorProps(color="FooterFillColor", style="FooterFillStyle")
        return self._props_internal_attributes
