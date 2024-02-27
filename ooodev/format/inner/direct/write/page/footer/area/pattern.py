# region Import
from __future__ import annotations
import uno

from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.common.props.area_pattern_props import AreaPatternProps
from ooodev.format.inner.direct.write.page.header.area.pattern import Pattern as HeaderPattern

# endregion Import


class Pattern(HeaderPattern):
    """
    Pattern fill for footer area.

    .. versionadded:: 0.9.2
    """

    # region Properties
    @property
    def _props(self) -> AreaPatternProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = AreaPatternProps(
                style="FooterFillStyle",
                name="FooterFillBitmapName",
                tile="FooterFillBitmapTile",
                stretch="FooterFillBitmapStretch",
                bitmap="FooterFillBitmap",
            )
        return self._props_internal_attributes

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.FOOTER
        return self._format_kind_prop

    # endregion Properties
