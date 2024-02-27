from __future__ import annotations
import uno
from ooodev.format.calc.style.page.kind.calc_style_page_kind import CalcStylePageKind
from ooodev.format.inner.common.props.hf_props import HfProps
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.modify.calc.page.header.header import Header


class Footer(Header):
    """
    Page Footer Settings

    .. seealso::

        - :ref:`help_calc_format_modify_page_footer_footer`

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        on: bool | None = None,
        shared: bool | None = None,
        shared_first: bool | None = None,
        margin_left: float | None = None,
        margin_right: float | None = None,
        spacing: float | None = None,
        height: float | None = None,
        height_auto: bool | None = None,
        style_name: CalcStylePageKind | str = CalcStylePageKind.DEFAULT,
        style_family: str = "PageStyles",
    ) -> None:
        """
        Constructor

        Args:
            on (bool | None, optional): Specifics if Footer is on.
            shared (bool | None, optional): Specifies if same contents left and right.
            shared_first (bool | None, optional): Specifies if same contents on first page.
            margin_left (float | None, optional): Specifies Left Margin in ``mm`` units.
            margin_right (float | None, optional): Specifies Right Margin in ``mm`` units.
            spacing (float | None, optional): Specifies Spacing in ``mm`` units.
            height (float | None, optional): Specifies Height in ``mm`` units.
            height_auto (bool | None, optional): Specifies if auto-fit height is used.
            style_name (CalcStylePageKind, str, optional): Specifies the Page Style that instance applies to.
                Default is Default Page Style.
            style_family (str, optional): Style family. Default ``PageStyles``.

        Returns:
            None:

        See Also:
            - :ref:`help_calc_format_modify_page_footer_footer`
        """
        super().__init__(
            on=on,
            shared=shared,
            shared_first=shared_first,
            margin_left=margin_left,
            margin_right=margin_right,
            spacing=spacing,
            height=height,
            height_auto=height_auto,
            style_name=style_name,
            style_family=style_family,
        )

    # region Internal Methods
    def _get_inner_props(self) -> HfProps:
        return HfProps(
            on="FooterIsOn",
            shared="FooterIsShared",
            shared_first="FirstPageFooterIsShared",
            margin_left="FooterLeftMargin",
            margin_right="FooterRightMargin",
            spacing="FooterBodyDistance",
            spacing_dyn="FooterDynamic",
            height="FooterHeight",
            height_auto="FooterIsDynamicHeight",
        )

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.DOC | FormatKind.STYLE | FormatKind.FOOTER
        return self._format_kind_prop

    # endregion Internal Methods
