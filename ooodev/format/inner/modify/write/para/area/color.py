# region Import
from __future__ import annotations
from typing import cast
import uno
from ooodev.format.writer.style.para.kind import StyleParaKind as StyleParaKind
from ooodev.utils import color as mColor
from ooodev.format.inner.direct.write.para.area.color import Color as InnerColor
from ..para_style_base_multi import ParaStyleBaseMulti

# endregion Import


class Color(ParaStyleBaseMulti):
    """
    Paragraph Style Fill Coloring

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        color: mColor.Color = -1,
        style_name: StyleParaKind | str = StyleParaKind.STANDARD,
        style_family: str = "ParagraphStyles",
    ) -> None:
        """
        Constructor

        Args:
            color (Color, optional): Fill Color.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to.
                Default is Default Paragraph Style.
            style_family (str, optional): Style family. Default ``ParagraphStyles``.

        Returns:
            None:
        """

        direct = InnerColor(color=color)
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
    ) -> Color:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Document Object.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to.
                Default is Default Paragraph Style.
            style_family (str, optional): Style family. Default ``ParagraphStyles``.

        Returns:
            Color: ``Color`` instance from document properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerColor.from_obj(inst.get_style_props(doc))
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
    def prop_inner(self) -> InnerColor:
        """Gets Inner Color instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(InnerColor, self._get_style_inst("direct"))
        return self._direct_inner
