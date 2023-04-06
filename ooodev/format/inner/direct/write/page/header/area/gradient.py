# region Import
from __future__ import annotations
from typing import Tuple
import uno
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.common.props.area_gradient_props import AreaGradientProps
from ooodev.format.inner.direct.write.fill.area.gradient import Gradient as InnerGradient

# endregion Import


class Gradient(InnerGradient):
    """
    Gradient of the header area of a page style

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
    def _props(self) -> AreaGradientProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = AreaGradientProps(
                style="HeaderFillStyle",
                step_count="HeaderFillGradientStepCount",
                name="HeaderFillGradientName",
                grad_prop_name="HeaderFillGradient",
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
