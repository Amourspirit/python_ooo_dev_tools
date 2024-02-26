from __future__ import annotations
import uno
from ooodev.format.inner.common.props.fill_color_props import FillColorProps
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.modify.write.page.header.area.color import Color as HeaderColor
from ooodev.format.writer.style.page.kind.writer_style_page_kind import WriterStylePageKind
from ooodev.utils import color as mColor
from ooodev.utils.color import StandardColor


class Color(HeaderColor):
    """
    Page Footer Color

    .. seealso::

        - :ref:`help_writer_format_modify_page_footer_area`

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        color: mColor.Color = StandardColor.AUTO_COLOR,
        style_name: WriterStylePageKind | str = WriterStylePageKind.STANDARD,
        style_family: str = "PageStyles",
    ) -> None:
        """
        Constructor

        Args:
            color (:py:data:`~.utils.color.Color`, optional): Footer Area Color.
            style_name (WriterStylePageKind, str, optional): Specifies the Page Style that instance applies to.
                Default is Default Page Style.
            style_family (str, optional): Style family. Default ``PageStyles``.

        Returns:
            None:

        See Also:
            - :ref:`help_writer_format_modify_page_footer_area`
        """
        super().__init__(color=color, style_name=style_name, style_family=style_family)

    # region internal methods
    def _get_inner_props(self) -> FillColorProps:
        return FillColorProps(color="FooterFillColor", style="FooterFillStyle")

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
