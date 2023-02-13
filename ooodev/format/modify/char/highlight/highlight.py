from __future__ import annotations
from typing import cast
import uno
from ....writer.style.char.kind.style_char_kind import StyleCharKind as StyleCharKind
from ..char_style_base_multi import CharStyleBaseMulti
from .....utils.color import Color
from ....direct.char.highlight.highlight import Highlight as DirectHighlight


class Highlight(CharStyleBaseMulti):
    """
    Character Style Highlight.

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        color: Color = -1,
        style_name: StyleCharKind | str = StyleCharKind.STANDARD,
        style_family: str = "CharacterStyles",
    ) -> None:
        """
        Constructor

        Args:
            color (Color, optional): Highlight Color
            style_name (StyleParaKind, str, optional): Specifies the Character Style that instance applies to. Deftult is Default Character Style.
            style_family (str, optional): Style family. Defatult ``CharacterStyles``.

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
        style_name: StyleCharKind | str = StyleCharKind.STANDARD,
        style_family: str = "CharacterStyles",
    ) -> Highlight:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Documnet Object.
            style_name (StyleCharKind, str, optional): Specifies the Character Style that instance applies to. Deftult is Default Character Style.
            style_family (str, optional): Style family. Defatult ``CharacterStyles``.

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
    def prop_style_name(self, value: str | StyleCharKind):
        self._style_name = str(value)

    @property
    def prop_inner(self) -> DirectHighlight:
        """Gets Inner Highlight instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(DirectHighlight, self._get_style_inst("direct"))
        return self._direct_inner
