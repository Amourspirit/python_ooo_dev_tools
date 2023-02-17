from __future__ import annotations
from typing import Tuple, cast
import uno
from ....writer.style.page.kind.style_page_kind import StylePageKind as StylePageKind
from ..page_style_base_multi import PageStyleBaseMulti
from ....direct.common.abstract.abstract_hf import AbstractHF
from ....direct.common.props.hf_props import HfProps


class PageFooter(AbstractHF):
    """
    Page Footer Settings

    .. versionadded:: 0.9.0
    """

    def _supported_services(self) -> Tuple[str, ...]:
        return ("com.sun.star.style.PageProperties", "com.sun.star.style.PageStyle")

    @property
    def _props(self) -> HfProps:
        try:
            return self._props_header
        except AttributeError:
            self._props_header = HfProps(
                on="FooterIsOn",
                shared="FooterIsShared",
                shared_first="FirstIsShared",
                margin_left="FooterLeftMargin",
                margin_right="FooterRightMargin",
                spacing="FooterBodyDistance",
                spacing_dyn="FooterDynamicSpacing",
                height="FooterHeight",
                height_auto="FooterIsDynamicHeight",
            )
        return self._props_header


class Footer(PageStyleBaseMulti):
    """
    Page Header Settings

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
        spacing_dyn: bool | None = None,
        height: float | None = None,
        height_auto: bool | None = None,
        style_name: StylePageKind | str = StylePageKind.STANDARD,
        style_family: str = "PageStyles",
    ) -> None:
        """
        Constructor

        Args:
            on (bool | None, optional): Specifices if Footer is on.
            shared (bool | None, optional): Specifies if same contents left and right.
            shared_first (bool | None, optional): Specifies if same contents on first page.
            margin_left (float | None, optional): Specifies Left Margin in ``mm`` units.
            margin_right (float | None, optional): Specifies Right Margin in ``mm`` units.
            spacing (float | None, optional): Specifies Spacing in ``mm`` units.
            spacing_dyn (bool | None, optional): Specifies if if dynamic spacing is used.
            height (float | None, optional): Specifies Height in ``mm`` units.
            height_auto (bool | None, optional): Specifies if auto fit height is used.
            style_name (StyleParaKind, str, optional): Specifies the Page Style that instance applies to. Deftult is Default Page Style.
            style_family (str, optional): Style family. Defatult ``PageStyles``.

        Returns:
            None:
        """

        direct = PageFooter(
            on=on,
            shared=shared,
            shared_first=shared_first,
            margin_left=margin_left,
            margin_right=margin_right,
            spacing=spacing,
            spacing_dyn=spacing_dyn,
            height=height,
            height_auto=height_auto,
        )
        super().__init__()
        self._style_name = str(style_name)
        self._style_family_name = style_family
        self._set_style("direct", direct, *direct.get_attrs())

    @classmethod
    def from_style(
        cls,
        doc: object,
        style_name: StylePageKind | str = StylePageKind.STANDARD,
        style_family: str = "PageStyles",
    ) -> Footer:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Documnet Object.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to. Deftult is Default Paragraph Style.
            style_family (str, optional): Style family. Defatult ``PageStyles``.

        Returns:
            Footer: ``Footer`` instance from document properties.
        """
        inst = super(Footer, cls).__new__(cls)
        inst.__init__(style_name=style_name, style_family=style_family)
        direct = PageFooter.from_obj(inst.get_style_props(doc))
        inst._set_style("direct", direct, *direct.get_attrs())
        return inst

    @property
    def prop_style_name(self) -> str:
        """Gets/Sets property Style Name"""
        return self._style_name

    @prop_style_name.setter
    def prop_style_name(self, value: str | StylePageKind):
        self._style_name = str(value)

    @property
    def prop_inner(self) -> PageFooter:
        """Gets Inner Footer instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(PageFooter, self._get_style_inst("direct"))
        return self._direct_inner