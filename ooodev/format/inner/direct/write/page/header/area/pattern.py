# region Import
from __future__ import annotations
from typing import Tuple
import uno

from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.common.props.area_pattern_props import AreaPatternProps
from ooodev.format.inner.direct.write.fill.area.pattern import Pattern as InnerPattern

# endregion Import


class Pattern(InnerPattern):
    """
    Pattern fill for header area.

    .. versionadded:: 0.9.2
    """

    # region Methods
    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = (
                "com.sun.star.style.PageProperties",
                "com.sun.star.style.PageStyle",
            )
        return self._supported_services_values

    # endregion Methods

    # region Properties
    @property
    def _props(self) -> AreaPatternProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = AreaPatternProps(
                style="HeaderFillStyle",
                name="HeaderFillBitmapName",
                tile="HeaderFillBitmapTile",
                stretch="HeaderFillBitmapStretch",
                bitmap="HeaderFillBitmap",
            )
        return self._props_internal_attributes

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.HEADER
        return self._format_kind_prop

    # endregion Properties
