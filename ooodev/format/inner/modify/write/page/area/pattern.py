# region Imports
from __future__ import annotations
from typing import cast
import uno
from com.sun.star.awt import XBitmap

from ooodev.format.inner.direct.write.fill.area.pattern import Pattern as InnerPattern
from ooodev.format.inner.modify.write.page.page_style_base_multi import PageStyleBaseMulti
from ooodev.format.inner.preset.preset_pattern import PresetPatternKind
from ooodev.format.writer.style.page.kind.writer_style_page_kind import WriterStylePageKind

# endregion Imports


class Pattern(PageStyleBaseMulti):
    """
    Page Style Pattern

    .. seealso::

        - :ref:`help_writer_format_modify_page_area`

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
        style_name: WriterStylePageKind | str = WriterStylePageKind.STANDARD,
        style_family: str = "PageStyles",
    ) -> None:
        """
        Constructor

        Args:
            bitmap (XBitmap, optional): Bitmap instance. If ``name`` is not already in the Bitmap Table then
                this property is required.
            name (str, optional): Specifies the name of the pattern. This is also the name that is used to store
                bitmap in LibreOffice Bitmap Table.
            tile (bool, optional): Specified if bitmap is tiled. Defaults to ``True``.
            stretch (bool, optional): Specifies if bitmap is stretched. Defaults to ``False``.
            auto_name (bool, optional): Specifies if ``name`` is ensured to be unique. Defaults to ``False``.
            style_name (StyleParaKind, str, optional): Specifies the Page Style that instance applies to.
                Default is Default Page Style.
            style_family (str, optional): Style family. Default ``PageStyles``.

        Returns:
            None:

        See Also:
            - :ref:`help_writer_format_modify_page_area`
        """

        direct = InnerPattern(bitmap=bitmap, name=name, tile=tile, stretch=stretch, auto_name=auto_name)
        super().__init__()
        self._style_name = str(style_name)
        self._style_family_name = style_family
        self._set_style("direct", direct, *direct.get_attrs())

    @classmethod
    def from_style(
        cls,
        doc: object,
        style_name: WriterStylePageKind | str = WriterStylePageKind.STANDARD,
        style_family: str = "PageStyles",
    ) -> Pattern:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Document Object.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to.
                Default is Default Paragraph Style.
            style_family (str, optional): Style family. Default ``PageStyles``.

        Returns:
            Pattern: ``Pattern`` instance from document properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerPattern.from_obj(inst.get_style_props(doc))
        inst._set_style("direct", direct, *direct.get_attrs())
        return inst

    @classmethod
    def from_preset(
        cls,
        preset: PresetPatternKind,
        style_name: WriterStylePageKind | str = WriterStylePageKind.STANDARD,
        style_family: str = "PageStyles",
    ) -> Pattern:
        """
        Gets an instance from a preset.

        Args:
            preset (PresetPatternKind): Preset.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to.
                Default is Default Paragraph Style.
            style_family (str, optional): Style family. Default ``PageStyles``.

        Returns:
            Pattern: ``Pattern`` instance from preset.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerPattern.from_preset(preset=preset)
        inst._set_style("direct", direct, *direct.get_attrs())
        return inst

    @property
    def prop_style_name(self) -> str:
        """Gets/Sets property Style Name"""
        return self._style_name

    @prop_style_name.setter
    def prop_style_name(self, value: str | WriterStylePageKind):
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
