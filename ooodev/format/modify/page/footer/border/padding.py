from __future__ import annotations
import uno
from .....writer.style.page.kind.writer_style_page_kind import WriterStylePageKind as WriterStylePageKind
from .....direct.common.props.border_props import BorderProps
from ...header.border.padding import Padding as HeaderPadding


class Padding(HeaderPadding):
    """
    Page Style Footer Border Padding.

    .. versionadded:: 0.9.0
    """

    # region Internal Methods
    def _get_inner_props(self) -> BorderProps:
        return BorderProps(
            left="FooterLeftBorderDistance",
            top="FooterTopBorderDistance",
            right="FooterRightBorderDistance",
            bottom="FooterBottomBorderDistance",
        )

    # endregion Internal Methods
