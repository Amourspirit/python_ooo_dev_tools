from __future__ import annotations

from ooodev.format.inner.common.props.area_img_props import AreaImgProps
from ...header.area.img import Img as HeaderImg


class Img(HeaderImg):
    """
    Page Footer Background Image

    .. versionadded:: 0.9.0
    """

    # region internal methods
    def _get_inner_props(self) -> AreaImgProps:
        return AreaImgProps(
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

    # endregion internal methods
