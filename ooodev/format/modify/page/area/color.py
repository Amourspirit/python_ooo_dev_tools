from __future__ import annotations
from typing import cast
import uno
from .....utils import color as mColor
from ....writer.style.page.kind.style_page_kind import StylePageKind as StylePageKind
from ..page_style_base_multi import PageStyleBaseMulti

from ....direct.para.area.color import Color as InnerColor


class Color(PageStyleBaseMulti):
    """
    Page Style Color

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        color: mColor.Color = -1,
        style_name: StylePageKind | str = StylePageKind.STANDARD,
        style_family: str = "PageStyles",
    ) -> None:
        """
        Constructor

        Args:
            color (Color, optional): FillColor Color
            style_name (StyleParaKind, str, optional): Specifies the Page Style that instance applies to. Deftult is Default Page Style.
            style_family (str, optional): Style family. Defatult ``PageStyles``.

        Returns:
            None:
        """

        direct = InnerColor(color=color)
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
    ) -> Color:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Documnet Object.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to. Deftult is Default Paragraph Style.
            style_family (str, optional): Style family. Defatult ``PageStyles``.

        Returns:
            Color: ``Color`` instance from document properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerColor.from_obj(inst.get_style_props(doc))
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
    def prop_inner(self) -> InnerColor:
        """Gets/Sets Inner Color instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(InnerColor, self._get_style_inst("direct"))
        return self._direct_inner

    @prop_inner.setter
    def prop_inner(self, value: InnerColor) -> None:
        if not isinstance(value, InnerColor):
            raise TypeError(f'Expected type of InnerColor, got "{type(value).__name__}"')
        self._del_attribs("_direct_inner")
        self._set_style("direct", value, *value.get_attrs())
