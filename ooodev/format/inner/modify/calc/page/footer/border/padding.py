# region Import
from __future__ import annotations

from ooodev.format.calc.style.page.kind import CalcStylePageKind as CalcStylePageKind
from ooodev.format.inner.common.props.border_props import BorderProps
from ooodev.format.inner.kind.format_kind import FormatKind
from ...header.border.padding import Padding as HeaderPadding

# endregion Import


class Padding(HeaderPadding):
    """
    Page Style Footer Border Padding.

    .. seealso::

        - :ref:`help_calc_format_modify_page_footer_borders`

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        left: float | None = None,
        right: float | None = None,
        top: float | None = None,
        bottom: float | None = None,
        padding_all: float | None = None,
        style_name: CalcStylePageKind | str = CalcStylePageKind.DEFAULT,
        style_family: str = "PageStyles",
    ) -> None:
        """
        Constructor

        Args:
            left (float, UnitT, optional): Left (in ``mm`` units) or :ref:`proto_unit_obj`.
            right (float, UnitT, optional): Right (in ``mm`` units)  or :ref:`proto_unit_obj`.
            top (float, UnitT, optional): Top (in ``mm`` units)  or :ref:`proto_unit_obj`.
            bottom (float, UnitT,  optional): Bottom (in ``mm`` units)  or :ref:`proto_unit_obj`.
            all (float, UnitT, optional): Left, right, top, bottom (in ``mm`` units)  or :ref:`proto_unit_obj`.
                If argument is present then ``left``, ``right``, ``top``, and ``bottom`` arguments are ignored.
            style_name (CalcStylePageKind, str, optional): Specifies the Page Style that instance applies to.
                Default is Default Page Style.
            style_family (str, optional): Style family. Default ``PageStyles``.

        Returns:
            None:

        See Also:
            - :ref:`help_calc_format_modify_page_footer_borders`
        """
        super().__init__(
            left=left,
            right=right,
            top=top,
            bottom=bottom,
            padding_all=padding_all,
            style_name=style_name,
            style_family=style_family,
        )

    # region overrides
    def _get_inner_props(self) -> BorderProps:
        return BorderProps(
            left="FooterLeftBorderDistance",
            top="FooterTopBorderDistance",
            right="FooterRightBorderDistance",
            bottom="FooterBottomBorderDistance",
        )

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.DOC | FormatKind.STYLE | FormatKind.FOOTER
        return self._format_kind_prop


# endregion overrides
