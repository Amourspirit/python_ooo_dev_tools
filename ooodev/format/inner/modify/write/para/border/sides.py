# region Import
from __future__ import annotations
from typing import Any, cast
import uno

from ooodev.format.writer.style.para.kind.style_para_kind import StyleParaKind
from ooodev.format.inner.direct.structs.side import Side
from ooodev.format.inner.direct.write.para.border.sides import Sides as InnerSides
from ooodev.format.inner.modify.write.para.para_style_base_multi import ParaStyleBaseMulti

# endregion Import


class Sides(ParaStyleBaseMulti):
    """
    Paragraph Style Sides (lines)

    .. seealso::

        - :ref:`help_writer_format_modify_para_borders`

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        left: Side | None = None,
        right: Side | None = None,
        top: Side | None = None,
        bottom: Side | None = None,
        all: Side | None = None,
        style_name: StyleParaKind | str = StyleParaKind.STANDARD,
        style_family: str = "ParagraphStyles",
    ) -> None:
        """
        Constructor

        Args:
            left (Side, optional): Determines the line style at the left edge.
            right (Side, optional): Determines the line style at the right edge.
            top (Side, optional): Determines the line style at the top edge.
            bottom (Side, optional): Determines the line style at the bottom edge.
            all (Side, optional): Determines the line style at the top, bottom, left, right edges.
                If this argument has a value then arguments ``top``, ``bottom``, ``left``, ``right`` are ignored
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to.
                Default is Default Paragraph Style.
            style_family (str, optional): Style family. Default ``ParagraphStyles``.

        Returns:
            None:

        See Also:
            - :ref:`help_writer_format_modify_para_borders`
        """

        direct = InnerSides(left=left, right=right, top=top, bottom=bottom, all=all)
        super().__init__()
        self._style_name = str(style_name)
        self._style_family_name = style_family
        self._set_style("direct", direct, *direct.get_attrs())

    @classmethod
    def from_style(
        cls,
        doc: Any,
        style_name: StyleParaKind | str = StyleParaKind.STANDARD,
        style_family: str = "ParagraphStyles",
    ) -> Sides:
        """
        Gets instance from Document.

        Args:
            doc (Any): UNO Document Object.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to.
                Default is Default Paragraph Style.
            style_family (str, optional): Style family. Default ``ParagraphStyles``.

        Returns:
            Sides: ``Sides`` instance from document properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerSides.from_obj(inst.get_style_props(doc))
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
    def prop_inner(self) -> InnerSides:
        """Gets/Sets Inner Padding instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(InnerSides, self._get_style_inst("direct"))
        return self._direct_inner

    @prop_inner.setter
    def prop_inner(self, value: InnerSides) -> None:
        if not isinstance(value, InnerSides):
            raise TypeError(f'Expected type of InnerSides, got "{type(value).__name__}"')
        self._del_attribs("_direct_inner")
        self._set_style("direct", value, *value.get_attrs())
