# region Import
from __future__ import annotations
from typing import Any, cast
import uno

from ooodev.format.inner.direct.write.char.highlight.highlight import Highlight as InnerHighlight
from ooodev.format.inner.modify.write.para.para_style_base_multi import ParaStyleBaseMulti
from ooodev.format.writer.style.para.kind.style_para_kind import StyleParaKind
from ooodev.utils.color import Color
from ooodev.utils.color import StandardColor

# endregion Import


class Highlight(ParaStyleBaseMulti):
    """
    Paragraph Style Highlight

    .. seealso::

        - :ref:`help_writer_format_modify_para_highlight`

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        color: Color = StandardColor.AUTO_COLOR,
        style_name: StyleParaKind | str = StyleParaKind.STANDARD,
        style_family: str = "ParagraphStyles",
    ) -> None:
        """
        Constructor

        Args:
            color (:py:data:`~.utils.color.Color`, optional): Highlight Color. A value of ``-1`` Set color to Transparent.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to.
                Default is Default Paragraph Style.
            style_family (str, optional): Style family. Default ``ParagraphStyles``.

        Returns:
            None:

        See Also:
            - :ref:`help_writer_format_modify_para_highlight`
        """

        direct = InnerHighlight(color=color)
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
    ) -> Highlight:
        """
        Gets instance from Document.

        Args:
            doc (Any): UNO Document Object.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to.
                Default is Default Paragraph Style.
            style_family (str, optional): Style family. Default ``ParagraphStyles``.

        Returns:
            Highlight: ``Highlight`` instance from document properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerHighlight.from_obj(inst.get_style_props(doc))
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
    def prop_inner(self) -> InnerHighlight:
        """Gets/Sets Inner Highlight instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(InnerHighlight, self._get_style_inst("direct"))
        return self._direct_inner

    @prop_inner.setter
    def prop_inner(self, value: InnerHighlight) -> None:
        if not isinstance(value, InnerHighlight):
            raise TypeError(f'Expected type of InnerHighlight, got "{type(value).__name__}"')
        self._del_attribs("_direct_inner")
        self._set_style("direct", value, *value.get_attrs())
