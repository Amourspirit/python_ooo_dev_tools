from __future__ import annotations
import uno
from ooodev.format.inner.common.props.area_img_props import AreaImgProps
from ooodev.format.inner.kind.format_kind import FormatKind
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
