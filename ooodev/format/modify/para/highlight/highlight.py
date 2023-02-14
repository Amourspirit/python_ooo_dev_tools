from __future__ import annotations
from typing import cast
import uno
from ....writer.style.para.kind import StyleParaKind as StyleParaKind
from ..para_style_base_multi import ParaStyleBaseMulti
from ....direct.char.highlight.highlight import Highlight as DirectHighlight
from .....utils.color import Color


class Highlight(ParaStyleBaseMulti):
    """
    Paragraph Style Highlight

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        color: Color = -1,
        style_name: StyleParaKind | str = StyleParaKind.STANDARD,
        style_family: str = "ParagraphStyles",
    ) -> None:
        """
        Constructor

        Args:
            color (Color, optional): Highlight Color. A value of ``-1`` Set color to Transparent.
            auto (bool, optional): Determines if the first line should be indented automatically.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to. Deftult is Default Paragraph Style.
            style_family (str, optional): Style family. Defatult ``ParagraphStyles``.

        Returns:
            None:
        """

        direct = DirectHighlight(color=color)
        super().__init__()
        self._style_name = str(style_name)
        self._style_family_name = style_family
        self._set_style("direct", direct, *direct.get_attrs())

    @classmethod
    def from_style(
        cls,
        doc: object,
        style_name: StyleParaKind | str = StyleParaKind.STANDARD,
        style_family: str = "ParagraphStyles",
    ) -> Highlight:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Documnet Object.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to. Deftult is Default Paragraph Style.
            style_family (str, optional): Style family. Defatult ``ParagraphStyles``.

        Returns:
            Highlight: ``Highlight`` instance from document properties.
        """
        inst = super(Highlight, cls).__new__(cls)
        inst.__init__(style_name=style_name, style_family=style_family)
        direct = DirectHighlight.from_obj(inst.get_style_props(doc))
        inst._set_style("direct", direct, *direct.get_attrs())
        return inst

    @property
    def prop_style_name(self) -> str:
        """Gets/Sets property Style Name"""
        return self._style_name

    @prop_style_name.setter
    def prop_style_name(self, value: str | StyleParaKind):
        self._style_name = str(value)

    @property
    def prop_inner(self) -> DirectHighlight:
        """Gets Inner Highlight instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(DirectHighlight, self._get_style_inst("direct"))
        return self._direct_inner
