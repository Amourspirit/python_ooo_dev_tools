from __future__ import annotations
from typing import cast
import uno
from ....writer.style.para.kind import StyleParaKind as StyleParaKind
from ..para_style_base_multi import ParaStyleBaseMulti
from ....direct.para.indent_space.spacing import Spacing as DirectSpacing


class Spacing(ParaStyleBaseMulti):
    """
    Paragraph Style Spacing

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        above: float | None = None,
        below: float | None = None,
        style_no_space: bool | None = None,
        style_name: StyleParaKind | str = StyleParaKind.STANDARD,
        style_family: str = "ParagraphStyles",
    ) -> None:
        """
        Constructor

        Args:
            above (float, optional): Determines the top margin of the paragraph (in mm units).
            below (float, optional): Determines the bottom margin of the paragraph (in mm units).
            style_no_space (bool, optional): Do not add space between paragraphs of the same style.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to. Deftult is Default Paragraph Style.
            style_family (str, optional): Style family. Defatult ``ParagraphStyles``.

        Returns:
            None:
        """

        direct = DirectSpacing(above=above, below=below, style_no_space=style_no_space)
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
    ) -> Spacing:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Documnet Object.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to. Deftult is Default Paragraph Style.
            style_family (str, optional): Style family. Defatult ``ParagraphStyles``.

        Returns:
            Spacing: ``Spacing`` instance from document properties.
        """
        inst = super(Spacing, cls).__new__(cls)
        inst.__init__(style_name=style_name, style_family=style_family)
        direct = DirectSpacing.from_obj(inst.get_style_props(doc))
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
    def prop_inner(self) -> DirectSpacing:
        """Gets Inner Spacing instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(DirectSpacing, self._get_style_inst("direct"))
        return self._direct_inner
