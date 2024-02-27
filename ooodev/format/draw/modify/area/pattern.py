"""Module for Draw Style Fill Coloring."""

# region Import
from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno
from ooodev.format.writer.style.para.kind.style_para_kind import StyleParaKind as StyleParaKind
from ooodev.format.inner.modify.write.fill.fill_style_base_multi import FillStyleBaseMulti
from ooodev.format.draw.style.kind import DrawStyleFamilyKind
from ooodev.format.draw.style.lookup import FamilyGraphics
from ooodev.format.inner.direct.write.fill.area.pattern import Pattern as InnerPattern

from ooodev.format.inner.preset.preset_pattern import PresetPatternKind

if TYPE_CHECKING:
    from com.sun.star.awt import XBitmap

# endregion Import


class Pattern(FillStyleBaseMulti):
    """
    Draw Style Fill Pattern

    .. seealso::

        - :ref:`help_draw_format_modify_area_pattern`

    .. versionadded:: 0.17.9
    """

    def __init__(
        self,
        *,
        bitmap: XBitmap | None = None,
        name: str = "",
        tile: bool = True,
        stretch: bool = False,
        auto_name: bool = False,
        style_name: str = FamilyGraphics.DEFAULT_DRAWING_STYLE,
        style_family: str | DrawStyleFamilyKind = DrawStyleFamilyKind.GRAPHICS,
    ) -> None:
        """
        Constructor

        Args:
            bitmap (XBitmap, optional): Bitmap instance. If ``name`` is not already in the Bitmap Table then this property is required.
            name (str, optional): Specifies the name of the pattern. This is also the name that is used to store bitmap in LibreOffice Bitmap Table.
            tile (bool, optional): Specified if bitmap is tiled. Defaults to ``True``.
            stretch (bool, optional): Specifies if bitmap is stretched. Defaults to ``False``.
            auto_name (bool, optional): Specifies if ``name`` is ensured to be unique. Defaults to ``False``.
            style_name (FamilyGraphics, str, optional): Specifies the Style that instance applies to.
                Default is Default ``standard`` Style.
            style_family (str, DrawStyleFamilyKind, optional): Family Style. Defaults to ``graphics``.

        Returns:
            None:

        See Also:
            - :ref:`help_draw_format_modify_area_pattern`
        """

        direct = InnerPattern(
            bitmap=bitmap,
            name=name,
            tile=tile,
            stretch=stretch,
            auto_name=auto_name,
        )
        super().__init__()
        self._style_family_name = str(style_family)
        self._style_name = str(style_name)
        self._set_style("direct", direct)

    @classmethod
    def from_preset(
        cls,
        preset: PresetPatternKind,
        style_name: str = FamilyGraphics.DEFAULT_DRAWING_STYLE,
        style_family: str | DrawStyleFamilyKind = DrawStyleFamilyKind.GRAPHICS,
    ) -> Pattern:
        """
        Gets an instance from a preset

        Args:
            preset (PresetPatternKind): Preset

        Returns:
            Pattern: Instance from preset.
        """
        direct = InnerPattern.from_preset(preset)
        inst = cls(style_name=style_name, style_family=style_family)
        inst._set_style("direct", direct)
        return inst

    @classmethod
    def from_style(
        cls,
        doc: Any,
        style_name: FamilyGraphics | str = FamilyGraphics.DEFAULT_DRAWING_STYLE,
        style_family: str | DrawStyleFamilyKind = DrawStyleFamilyKind.GRAPHICS,
    ) -> Pattern:
        """
        Gets instance from Document.

        Args:
            doc (Any): UNO Document Object.
            style_name (FamilyGraphics, str, optional): Specifies the Style that instance applies to.
                Default is ``FamilyGraphics.DEFAULT_DRAWING_STYLE``.
            style_family (DrawStyleFamilyKind, str, optional): Style family. Default ``DrawStyleFamilyKind.GRAPHICS``.

        Returns:
            Pattern: ``Pattern`` instance from document properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerPattern.from_obj(inst.get_style_props(doc))
        inst._set_style("direct", direct)
        return inst

    @property
    def prop_style_name(self) -> str:
        """Gets/Sets property Style Name"""
        return self._style_name

    @prop_style_name.setter
    def prop_style_name(self, value: str | FamilyGraphics):
        self._style_name = str(value)

    @property
    def prop_inner(self) -> InnerPattern:
        """Gets Inner instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(InnerPattern, self._get_style_inst("direct"))
        return self._direct_inner
