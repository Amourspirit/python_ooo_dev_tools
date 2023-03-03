from __future__ import annotations
from typing import cast
import uno
from .....proto.unit_obj import UnitObj
from ....writer.style.para.kind import StyleParaKind as StyleParaKind
from ..para_style_base_multi import ParaStyleBaseMulti
from ....direct.para.indent_space.indent import Indent as InnerIndent


class Indent(ParaStyleBaseMulti):
    """
    Paragraph Style Indent

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        before: float | UnitObj | None = None,
        after: float | UnitObj | None = None,
        first: float | UnitObj | None = None,
        auto: bool | None = None,
        style_name: StyleParaKind | str = StyleParaKind.STANDARD,
        style_family: str = "ParagraphStyles",
    ) -> None:
        """
        Constructor

        Args:
            before (float, UnitObj, optional): Determines the left margin of the paragraph (in ``mm`` units) or :ref:`proto_unit_obj`.
            after (float, UnitObj, optional): Determines the right margin of the paragraph (in ``mm`` units) or :ref:`proto_unit_obj`.
            first (float, UnitObj, optional): specifies the indent for the first line (in ``mm`` units) or :ref:`proto_unit_obj`.
            auto (bool, optional): Determines if the first line should be indented automatically.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to. Deftult is Default Paragraph Style.
            style_family (str, optional): Style family. Defatult ``ParagraphStyles``.

        Returns:
            None:
        """

        direct = InnerIndent(before=before, after=after, first=first, auto=auto)
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
    ) -> Indent:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Documnet Object.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to. Deftult is Default Paragraph Style.
            style_family (str, optional): Style family. Defatult ``ParagraphStyles``.

        Returns:
            Indent: ``Indent`` instance from document properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerIndent.from_obj(inst.get_style_props(doc))
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
    def prop_inner(self) -> InnerIndent:
        """Gets Inner Indent instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(InnerIndent, self._get_style_inst("direct"))
        return self._direct_inner

    @prop_inner.setter
    def prop_inner(self, value: InnerIndent) -> None:
        if not isinstance(value, InnerIndent):
            raise TypeError(f'Expected type of InnerIndent, got "{type(value).__name__}"')
        self._del_attribs("_direct_inner")
        self._set_style("direct", value, *value.get_attrs())
