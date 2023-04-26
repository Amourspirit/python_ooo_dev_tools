# region Import
from __future__ import annotations
from typing import cast
import uno
from ooodev.format.writer.style.para.kind import StyleParaKind as StyleParaKind
from ooodev.format.inner.direct.structs.side import Side as Side, LineSize as LineSize
from ooodev.format.inner.direct.write.para.border.shadow import Shadow
from ooodev.format.inner.direct.write.para.border.padding import Padding
from ooodev.format.inner.direct.write.para.border.borders import Borders as InnerBorders
from ..para_style_base_multi import ParaStyleBaseMulti

# endregion Import


class Borders(ParaStyleBaseMulti):
    """
    Paragraph Style Borders

    .. seealso::

        - :ref:`help_writer_format_modify_para_borders`

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
            border_side (Side, None, optional): Determines the line style at the top, bottom, left, right edges.
                If this argument has a value then arguments ``top``, ``bottom``, ``left``, ``right`` are ignored
            shadow (BorderShadow, None, optional): Character Shadow
            padding (BorderPadding, None, optional): Character padding
            merge (bool, None, optional): Merge with next paragraph
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to.
                Default is Default Paragraph Style.
            style_family (str, optional): Style family. Default ``ParagraphStyles``.

        Returns:
            None:

        See Also:
            - :ref:`help_writer_format_modify_para_borders`
        """

        direct = InnerBorders(
            right=right,
            left=left,
            top=top,
            bottom=bottom,
            all=border_side,
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
            doc (object): UNO Document Object.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to.
                Default is Default Paragraph Style.
            style_family (str, optional): Style family. Default ``ParagraphStyles``.

        Returns:
            Borders: ``Borders`` instance from document properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerBorders.from_obj(inst.get_style_props(doc))
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
    def prop_inner(self) -> InnerBorders:
        """Gets/Sets Inner Borders instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(InnerBorders, self._get_style_inst("direct"))
        return self._direct_inner

    @prop_inner.setter
    def prop_inner(self, value: InnerBorders) -> None:
        if not isinstance(value, InnerBorders):
            raise TypeError(f'Expected type of InnerBorders, got "{type(value).__name__}"')
        self._del_attribs("_direct_inner")
        self._set_style("direct", value, *value.get_attrs())
