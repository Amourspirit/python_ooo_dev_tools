from __future__ import annotations
import uno
from ooodev.format.inner.common.props.area_gradient_props import AreaGradientProps
from ooodev.format.inner.kind.format_kind import FormatKind
from ...header.area.gradient import Gradient as HeaderGradient


class Gradient(HeaderGradient):
    """
    Page Footer Gradient Color

    .. versionadded:: 0.9.0
    """

    # region internal methods
    def _get_inner_props(self) -> AreaGradientProps:
        return AreaGradientProps(
            style="FooterFillStyle",
            step_count="FooterFillGradientStepCount",
            name="FooterFillGradientName",
            grad_prop_name="FooterFillGradient",
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
