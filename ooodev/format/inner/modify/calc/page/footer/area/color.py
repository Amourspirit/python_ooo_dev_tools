# region Import
from __future__ import annotations
from ooodev.format.calc.style.page.kind import CalcStylePageKind as CalcStylePageKind
from ooodev.format.inner.common.props.fill_color_props import FillColorProps
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.utils import color as mColor
from ...header.area.color import InnerColor as InnerColor
from ...header.area.color import Color as HeaderColor

# endregion Import


class Color(HeaderColor):
    """
    Page Footer Color

    .. seealso::

        - :ref:`help_calc_format_modify_page_footer_background`

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        color: mColor.Color = -1,
        style_name: CalcStylePageKind | str = CalcStylePageKind.DEFAULT,
        style_family: str = "PageStyles",
    ) -> None:
        """
        Constructor

        Args:
            color (:py:data:`~.utils.color.Color`, optional): FillColor Color.
            style_name (CalcStylePageKind, str, optional): Specifies the Page Style that instance applies to.
                Default is Default Page Style.
            style_family (str, optional): Style family. Default ``PageStyles``.

        Returns:
            None:

        See Also:
            - :ref:`help_calc_format_modify_page_footer_background`
        """
        super().__init__(color=color, style_name=style_name, style_family=style_family)

    # region internal methods
    def _get_inner_props(self) -> FillColorProps:
        return FillColorProps(color="FooterBackgroundColor", style="", bg="FooterBackTransparent")

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.DOC | FormatKind.STYLE | FormatKind.FOOTER
        return self._format_kind_prop

    # endregion internal methods
