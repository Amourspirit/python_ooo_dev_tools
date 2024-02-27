from __future__ import annotations
import uno

from ooodev.format.writer.style.page.kind.writer_style_page_kind import WriterStylePageKind
from ooodev.format.inner.direct.structs.side import Side
from ooodev.format.inner.common.props.border_props import BorderProps
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.modify.write.page.header.border.sides import Sides as HeaderSides


class Sides(HeaderSides):
    """
    Page Footer Style Border Sides.

    .. seealso::

        - :ref:`help_writer_format_modify_page_footer_borders`

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        left: Side | None = None,
        right: Side | None = None,
        top: Side | None = None,
        bottom: Side | None = None,
        all: Side | None = None,
        style_name: WriterStylePageKind | str = WriterStylePageKind.STANDARD,
        style_family: str = "PageStyles",
    ) -> None:
        """
        Constructor

        Args:
            left (Side | None, optional): Determines the line style at the left edge.
            right (Side | None, optional): Determines the line style at the right edge.
            top (Side | None, optional): Determines the line style at the top edge.
            bottom (Side | None, optional): Determines the line style at the bottom edge.
            all (Side | None, optional): Determines the line style at the top, bottom, left, right edges.
                If this argument has a value then arguments ``top``, ``bottom``, ``left``, ``right`` are ignored
            style_name (WriterStylePageKind, str, optional): Specifies the Page Style that instance applies to.
                Default is Default Page Style.
            style_family (str, optional): Style family. Default ``PageStyles``.

        Returns:
            None:

        See Also:
            - :ref:`help_writer_format_modify_page_footer_borders`
        """
        super().__init__(
            left=left, right=right, top=top, bottom=bottom, all=all, style_name=style_name, style_family=style_family
        )

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
