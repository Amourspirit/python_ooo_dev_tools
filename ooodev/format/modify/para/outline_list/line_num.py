from __future__ import annotations
from typing import cast
import uno
from ....writer.style.para.kind import StyleParaKind as StyleParaKind
from ..para_style_base_multi import ParaStyleBaseMulti
from ....kind.format_kind import FormatKind
from ....direct.common.abstract.abstract_line_number import AbstractLineNumber, LineNumeProps


class DirectLineNum(AbstractLineNumber):
    @property
    def _props(self) -> LineNumeProps:
        try:
            return self._props_line_num
        except AttributeError:
            self._props_line_num = LineNumeProps(value="ParaLineNumberStartValue", count="ParaLineNumberCount")
        return self._props_line_num

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        return FormatKind.PARA


class LineNum(ParaStyleBaseMulti):
    """
    Paragraph Style Line Number

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        num_start: int = 0,
        style_name: StyleParaKind | str = StyleParaKind.STANDARD,
        style_family: str = "ParagraphStyles",
    ) -> None:
        """
        Constructor

        Args:
            before (float, optional): Determines the left margin of the paragraph (in ``mm`` units).
            after (float, optional): Determines the right margin of the paragraph (in ``mm`` units).
            first (float, optional): specifies the indent for the first line (in ``mm`` units).
            auto (bool, optional): Determines if the first line should be indented automatically.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to. Deftult is Default Paragraph Style.
            style_family (str, optional): Style family. Defatult ``ParagraphStyles``.

        Returns:
            None:
        """

        direct = DirectLineNum(num_start=num_start)
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
    ) -> LineNum:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Documnet Object.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to. Deftult is Default Paragraph Style.
            style_family (str, optional): Style family. Defatult ``ParagraphStyles``.

        Returns:
            LineNum: ``LineNum`` instance from document properties.
        """
        inst = super(LineNum, cls).__new__(cls)
        inst.__init__(style_name=style_name, style_family=style_family)
        direct = DirectLineNum.from_obj(inst.get_style_props(doc))
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
    def prop_inner(self) -> DirectLineNum:
        """Gets Inner Line Number instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(DirectLineNum, self._get_style_inst("direct"))
        return self._direct_inner