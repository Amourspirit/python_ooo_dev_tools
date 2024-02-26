# region Import
from __future__ import annotations
import uno

from ooodev.format.inner.common.props.area_img_props import AreaImgProps
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.direct.write.page.header.area.img import Img as HeaderImg

# endregion Import


class Img(HeaderImg):
    """
    Img style for footer area

    .. versionadded:: 0.9.2
    """

    # region Properties
    @property
    def _props(self) -> AreaImgProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = AreaImgProps(
                name="FooterFillBitmapName",
                style="FooterFillStyle",
                mode="FooterFillBitmapMode",
                point="FooterFillBitmapRectanglePoint",
                bitmap="FooterFillBitmap",
                offset_x="FooterFillBitmapOffsetX",
                offset_y="FooterFillBitmapOffsetY",
                pos_x="FooterFillBitmapPositionOffsetX",
                pos_y="FooterFillBitmapPositionOffsetY",
                size_x="FooterFillBitmapSizeX",
                size_y="FooterFillBitmapSizeY",
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
