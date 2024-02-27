# region Import
from __future__ import annotations
from typing import Any, cast
import uno
from com.sun.star.awt import XBitmap

from ooodev.format.inner.direct.write.para.area.pattern import Pattern as InnerPattern
from ooodev.format.inner.modify.write.para.para_style_base_multi import ParaStyleBaseMulti
from ooodev.format.inner.preset.preset_pattern import PresetPatternKind
from ooodev.format.writer.style.para.kind.style_para_kind import StyleParaKind

# endregion Import


class Pattern(ParaStyleBaseMulti):
    """
    Paragraph Style Gradient Color

    .. seealso::

        - :ref:`help_writer_format_modify_para_pattern`

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        bitmap: XBitmap | None = None,
        name: str = "",
        tile: bool = True,
        stretch: bool = False,
        auto_name: bool = False,
        style_name: StyleParaKind | str = StyleParaKind.STANDARD,
        style_family: str = "ParagraphStyles",
    ) -> None:
        """
        Constructor

        Args:
            bitmap (XBitmap, optional): Bitmap instance. If ``name`` is not already in the Bitmap Table then this
                property is required.
            name (str, optional): Specifies the name of the pattern. This is also the name that is used to store
                bitmap in LibreOffice Bitmap Table.
            tile (bool, optional): Specified if bitmap is tiled. Defaults to ``True``.
            stretch (bool, optional): Specifies if bitmap is stretched. Defaults to ``False``.
            auto_name (bool, optional): Specifies if ``name`` is ensured to be unique. Defaults to ``False``.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to.
                Default is Default Paragraph Style.
            style_family (str, optional): Style family. Default ``ParagraphStyles``.

        Returns:
            None:

        See Also:
            - :ref:`help_writer_format_modify_para_pattern`
        """

        direct = InnerPattern(bitmap=bitmap, name=name, tile=tile, stretch=stretch, auto_name=auto_name)
        super().__init__()
        self._style_name = str(style_name)
        self._style_family_name = style_family
        self._set_style("direct", direct, *direct.get_attrs())

    @classmethod
    def from_preset(
        cls,
        preset: PresetPatternKind,
        style_name: StyleParaKind | str = StyleParaKind.STANDARD,
        style_family: str = "ParagraphStyles",
    ) -> Pattern:
        """
        Gets instance from preset

        Args:
            preset (PresetPatternKind): Preset

        Returns:
            Hatch: Hatch from a preset.
        """
        direct = InnerPattern.from_preset(preset)
        inst = cls(style_name=style_name, style_family=style_family)
        inst._set_style("direct", direct, *direct.get_attrs())
        return inst

    @classmethod
    def from_style(
        cls,
        doc: Any,
        style_name: StyleParaKind | str = StyleParaKind.STANDARD,
        style_family: str = "ParagraphStyles",
    ) -> Pattern:
        """
        Gets instance from Document.

        Args:
            doc (Any): UNO Document Object.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to.
                Default is Default Paragraph Style.
            style_family (str, optional): Style family. Default ``ParagraphStyles``.

        Returns:
            Alignment: ``Alignment`` instance from document properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerPattern.from_obj(inst.get_style_props(doc))
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
    def prop_inner(self) -> InnerPattern:
        """Gets/Sets Inner Pattern instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(InnerPattern, self._get_style_inst("direct"))
        return self._direct_inner

    @prop_inner.setter
    def prop_inner(self, value: InnerPattern) -> None:
        if not isinstance(value, InnerPattern):
            raise TypeError(f'Expected type of InnerPattern, got "{type(value).__name__}"')
        self._del_attribs("_direct_inner")
        self._set_style("direct", value, *value.get_attrs())
