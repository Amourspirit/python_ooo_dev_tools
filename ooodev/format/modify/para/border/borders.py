from __future__ import annotations
from typing import cast
import uno
from ....writer.style.para.kind import StyleParaKind as StyleParaKind
from ..para_style_base_multi import ParaStyleBaseMulti
from ....direct.structs.side import Side as Side, SideFlags as SideFlags, LineSize as LineSize
from ....direct.para.border.shadow import Shadow as Shadow
from ....direct.para.border.padding import Padding as Padding
from ....direct.para.border.borders import Borders as DirectBorders


class Borders(ParaStyleBaseMulti):
    """
    Paragraph Style Borders

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        right: Side | None = None,
        left: Side | None = None,
        top: Side | None = None,
        bottom: Side | None = None,
        border_side: Side | None = None,
        shadow: Shadow | None = None,
        padding: Padding | None = None,
        merge: bool | None = None,
        style_name: StyleParaKind | str = StyleParaKind.STANDARD,
        style_family: str = "ParagraphStyles",
    ) -> None:
        """
        Constructor

        Args:
            left (Side, None, optional): Determines the line style at the left edge.
            right (Side, None, optional): Determines the line style at the right edge.
            top (Side, None, optional): Determines the line style at the top edge.
            bottom (Side, None, optional): Determines the line style at the bottom edge.
            border_side (Side, None, optional): Determines the line style at the top, bottom, left, right edges. If this argument has a value then arguments ``top``, ``bottom``, ``left``, ``right`` are ignored
            shadow (BorderShadow, None, optional): Character Shadow
            padding (BorderPadding, None, optional): Character padding
            merge (bool, None, optional): Merge with next paragraph
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to. Deftult is Default Paragraph Style.
            style_family (str, optional): Style family. Defatult ``ParagraphStyles``.

        Returns:
            None:
        """

        direct = DirectBorders(
            right=right,
            left=left,
            top=top,
            bottom=bottom,
            border_side=border_side,
            shadow=shadow,
            padding=padding,
            merge=merge,
        )
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
    ) -> Borders:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Documnet Object.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to. Deftult is Default Paragraph Style.
            style_family (str, optional): Style family. Defatult ``ParagraphStyles``.

        Returns:
            Borders: ``Borders`` instance from document properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = DirectBorders.from_obj(inst.get_style_props(doc))
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
    def prop_inner(self) -> DirectBorders:
        """Gets Inner Borders instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(DirectBorders, self._get_style_inst("direct"))
        return self._direct_inner
