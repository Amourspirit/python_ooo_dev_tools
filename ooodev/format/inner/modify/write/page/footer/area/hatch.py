from __future__ import annotations
import uno
from ooodev.format.inner.common.props.area_hatch_props import AreaHatchProps
from ooodev.format.inner.kind.format_kind import FormatKind
from ...header.area.hatch import Hatch as HeaderHatch


class Hatch(HeaderHatch):
    """
    Page Footer Hatch
    .. versionadded:: 0.9.0
    """

    # region internal methods
    def _get_inner_props(self) -> AreaHatchProps:
        return AreaHatchProps(
            color="FooterFillColor",
            style="FooterFillStyle",
            bg="FooterFillBackground",
            hatch_prop="FooterFillHatch",
        )

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
