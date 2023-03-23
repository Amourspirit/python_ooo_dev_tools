# region Import
from __future__ import annotations
from typing import cast
import uno
from ooo.dyn.style.break_type import BreakType as BreakType

from ooodev.format.writer.style.para.kind import StyleParaKind as StyleParaKind
from ooodev.format.inner.direct.write.para.text_flow.breaks import Breaks as InnerBreaks
from ..para_style_base_multi import ParaStyleBaseMulti

# endregion Import


class Breaks(ParaStyleBaseMulti):
    """
    Paragraph Style Breaks

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        type: BreakType | None = None,
        style: str | None = None,
        num: int | None = None,
        style_name: StyleParaKind | str = StyleParaKind.STANDARD,
        style_family: str = "ParagraphStyles",
    ) -> None:
        """
        Constructor

        Args:
            type (BreakType, optional): Break type.
            style (str, optional): Style to apply to break.
            num (int, optional): Page number to apply to break.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to.
                Default is Default Paragraph Style.
            style_family (str, optional): Style family. Default ``ParagraphStyles``.

        Returns:
            None:
        """

        direct = InnerBreaks(type=type, style=style, num=num)
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
    ) -> Breaks:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Document Object.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to.
                Default is Default Paragraph Style.
            style_family (str, optional): Style family. Default ``ParagraphStyles``.

        Returns:
            Breaks: ``Breaks`` instance from document properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerBreaks.from_obj(inst.get_style_props(doc))
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
    def prop_inner(self) -> InnerBreaks:
        """Gets/Sets Inner Breaks instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(InnerBreaks, self._get_style_inst("direct"))
        return self._direct_inner

    @prop_inner.setter
    def prop_inner(self, value: InnerBreaks) -> None:
        if not isinstance(value, InnerBreaks):
            raise TypeError(f'Expected type of InnerBreaks, got "{type(value).__name__}"')
        self._del_attribs("_direct_inner")
        self._set_style("direct", value, *value.get_attrs())
