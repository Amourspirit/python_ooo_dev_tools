# region Import
from __future__ import annotations
from typing import Tuple
import uno
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.common.props.area_gradient_props import AreaGradientProps
from ooodev.format.inner.direct.write.page.header.area.gradient import Gradient as HeaderGradient

# endregion Import


class Gradient(HeaderGradient):
    """
    Gradient of the header area of a page style

    .. versionadded:: 0.9.2
    """

    # region Properties
    @property
    def _props(self) -> AreaGradientProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = AreaGradientProps(
                style="FooterFillStyle",
                step_count="FooterFillGradientStepCount",
                name="FooterFillGradientName",
                grad_prop_name="FooterFillGradient",
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
