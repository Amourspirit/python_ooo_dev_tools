# region Import
from __future__ import annotations
from typing import cast
import uno
from ooo.dyn.style.page_style_layout import PageStyleLayout
from ooo.dyn.style.numbering_type import NumberingTypeEnum

from ooodev.format.calc.style.page.kind.calc_style_page_kind import CalcStylePageKind
from ooodev.format.inner.direct.calc.page.page.layout_settings import LayoutSettings as InnerLayoutSettings
from ooodev.format.inner.modify.calc.cell_style_base_multi import CellStyleBaseMulti
from ooodev.format.writer.style.para.kind.style_para_kind import StyleParaKind

# endregion Import


class LayoutSettings(CellStyleBaseMulti):
    """
    Page Layout Setting style

    .. seealso::

        - :ref:`help_calc_format_modify_page_page`

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        layout: PageStyleLayout | None = None,
        numbers: NumberingTypeEnum | None = None,
        align_hori: bool | None = None,
        align_vert: bool | None = None,
        style_name: CalcStylePageKind | str = CalcStylePageKind.DEFAULT,
        style_family: str = "PageStyles",
    ) -> None:
        """
        Constructor

        Args:
            layout (PageStyleLayout, optional): Specifies the layout of the page.
            numbers (NumberingTypeEnum, optional): Specifies the default numbering type for this page.
            align_hori: bool | None = None,
            align_vert: bool | None = None,
            style_name (CalcStylePageKind, str, optional): Specifies the Page Style that instance applies to.
                Default is Default Page Style.
            style_family (str, optional): Style family. Default ``PageStyles``.

        Returns:
            None:

        See Also:
            - :ref:`help_calc_format_modify_page_page`
        """

        direct = InnerLayoutSettings(layout=layout, numbers=numbers, align_hori=align_hori, align_vert=align_vert)
        super().__init__()
        self._style_name = str(style_name)
        self._style_family_name = style_family
        self._set_style("direct", direct, *direct.get_attrs())

    @classmethod
    def from_style(
        cls,
        doc: object,
        style_name: CalcStylePageKind | str = CalcStylePageKind.DEFAULT,
        style_family: str = "PageStyles",
    ) -> LayoutSettings:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Document Object.
            style_name (CalcStylePageKind, str, optional): Specifies the Paragraph Style that instance applies to.
                Default is Default Paragraph Style.
            style_family (str, optional): Style family. Default ``PageStyles``.

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
    def prop_style_name(self, value: str | CalcStylePageKind):
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
