# region Import
from __future__ import annotations
from typing import Tuple
import uno

from ooodev.format.inner.common.props.area_img_props import AreaImgProps
from ooodev.format.inner.direct.write.fill.area.img import Img as InnerImg
from ooodev.format.inner.kind.format_kind import FormatKind

# endregion Import


class Img(InnerImg):
    """
    Img style for header area

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
    def _props(self) -> AreaImgProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = AreaImgProps(
                name="HeaderFillBitmapName",
                style="HeaderFillStyle",
                mode="HeaderFillBitmapMode",
                point="HeaderFillBitmapRectanglePoint",
                bitmap="HeaderFillBitmap",
                offset_x="HeaderFillBitmapOffsetX",
                offset_y="HeaderFillBitmapOffsetY",
                pos_x="HeaderFillBitmapPositionOffsetX",
                pos_y="HeaderFillBitmapPositionOffsetY",
                size_x="HeaderFillBitmapSizeX",
                size_y="HeaderFillBitmapSizeY",
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
