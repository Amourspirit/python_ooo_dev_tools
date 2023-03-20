from __future__ import annotations
from typing import cast
import uno
from ....writer.style.page.kind import WriterStylePageKind as WriterStylePageKind
from ..page_style_base_multi import PageStyleBaseMulti
from ....writer.style.para.kind.style_para_kind import StyleParaKind as StyleParaKind
from ....direct.page.page.layout_settings import LayoutSettings as InnerLayoutSettings

from ooo.dyn.style.page_style_layout import PageStyleLayout as PageStyleLayout
from ooo.dyn.style.numbering_type import NumberingTypeEnum as NumberingTypeEnum


class LayoutSettings(PageStyleBaseMulti):
    """
    Page Layout Setting style

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        layout: PageStyleLayout | None = None,
        numbers: NumberingTypeEnum | None = None,
        ref_style: str | StyleParaKind | None = None,
        right_gutter: bool | None = None,
        style_name: WriterStylePageKind | str = WriterStylePageKind.STANDARD,
        style_family: str = "PageStyles",
    ) -> None:
        """
        Constructor

        Args:
            layout (PageStyleLayout, optional): Specifies the layout of the page.
            numbers (NumberingTypeEnum, optional): Specifies the default numbering type for this page.
            ref_style (str, StyleParaKind, optional): Specifies the name of the paragraph style that is used as reference of the register mode.
            right_gutter (bool, optional): Specifies that the page gutter shall be placed on the right side of the page.
            style_name (StyleParaKind, str, optional): Specifies the Page Style that instance applies to. Deftult is Default Page Style.
            style_family (str, optional): Style family. Defatult ``PageStyles``.

        Returns:
            None:
        """

        direct = InnerLayoutSettings(layout=layout, numbers=numbers, ref_style=ref_style, right_gutter=right_gutter)
        super().__init__()
        self._style_name = str(style_name)
        self._style_family_name = style_family
        self._set_style("direct", direct, *direct.get_attrs())

    @classmethod
    def from_style(
        cls,
        doc: object,
        style_name: WriterStylePageKind | str = WriterStylePageKind.STANDARD,
        style_family: str = "PageStyles",
    ) -> LayoutSettings:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Documnet Object.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to. Deftult is Default Paragraph Style.
            style_family (str, optional): Style family. Defatult ``PageStyles``.

        Returns:
            LayoutSettings: ``LayoutSettings`` instance from document properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerLayoutSettings.from_obj(inst.get_style_props(doc))
        inst._set_style("direct", direct, *direct.get_attrs())
        return inst

    @property
    def prop_style_name(self) -> str:
        """Gets/Sets property Style Name"""
        return self._style_name

    @prop_style_name.setter
    def prop_style_name(self, value: str | WriterStylePageKind):
        self._style_name = str(value)

    @property
    def prop_inner(self) -> InnerLayoutSettings:
        """Gets/Sets Inner Layout Settings instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(InnerLayoutSettings, self._get_style_inst("direct"))
        return self._direct_inner

    @prop_inner.setter
    def prop_inner(self, value: InnerLayoutSettings) -> None:
        if not isinstance(value, InnerLayoutSettings):
            raise TypeError(f'Expected type of InnerLayoutSettings, got "{type(value).__name__}"')
        self._del_attribs("_direct_inner")
        self._set_style("direct", value, *value.get_attrs())
