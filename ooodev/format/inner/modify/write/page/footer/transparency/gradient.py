from __future__ import annotations
import uno
from ooodev.format.inner.common.props.transparent_gradient_props import TransparentGradientProps
from ooodev.format.inner.kind.format_kind import FormatKind
from ...header.transparency.gradient import Gradient as HeaderGradient


class Gradient(HeaderGradient):
    """
    Page Footer Transparent Gradient

    .. versionadded:: 0.9.0
    """

    # region Internal Methods
    def _get_inner_props(self) -> TransparentGradientProps:
        return TransparentGradientProps(
            transparence="FooterFillTransparence",
            name="FooterFillTransparenceGradientName",
            struct_prop="FooterFillTransparenceGradient",
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
