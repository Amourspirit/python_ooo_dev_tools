from __future__ import annotations
import uno
from ooodev.format.inner.common.props.border_props import BorderProps
from ooodev.format.inner.kind.format_kind import FormatKind
from ...header.border.sides import Sides as HeaderSides


class Sides(HeaderSides):
    """
    Page Footer Style Border Sides.

    .. versionadded:: 0.9.0
    """

    # region Internal Methods
    def _get_inner_props(self) -> BorderProps:
        return BorderProps(
            left="FooterLeftBorder", top="FooterTopBorder", right="FooterRightBorder", bottom="FooterBottomBorder"
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
