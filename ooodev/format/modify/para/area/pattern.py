from __future__ import annotations
from typing import cast
import uno
from .....utils.color import Color
from .....utils.data_type.angle import Angle as Angle
from .....utils.data_type.color_range import ColorRange as ColorRange
from .....utils.data_type.intensity import Intensity as Intensity
from .....utils.data_type.intensity_range import IntensityRange as IntensityRange
from .....utils.data_type.offset import Offset as Offset
from ....writer.style.para.kind import StyleParaKind as StyleParaKind
from ....preset import preset_pattern
from ....preset.preset_pattern import PresetPatternKind as PresetPatternKind
from ..para_style_base_multi import ParaStyleBaseMulti

from ....direct.para.area.pattern import Pattern as DirectPattern

from ooo.dyn.drawing.hatch_style import HatchStyle as HatchStyle

from com.sun.star.awt import XBitmap


class Pattern(ParaStyleBaseMulti):
    """
    Paragraph Style Gradient Color

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
            bitmap (XBitmap, optional): Bitmap instance. If ``name`` is not already in the Bitmap Table then this property is requied.
            name (str, optional): Specifies the name of the pattern. This is also the name that is used to store bitmap in LibreOffice Bitmap Table.
            tile (bool, optional): Specified if bitmap is tiled. Defaults to ``True``.
            stretch (bool, optional): Specifies if bitmap is stretched. Defaults to ``False``.
            auto_name (bool, optional): Specifies if ``name`` is ensured to be unique. Defaults to ``False``.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to. Deftult is Default Paragraph Style.
            style_family (str, optional): Style family. Defatult ``ParagraphStyles``.

        Returns:
            None:
        """

        direct = DirectPattern(bitmap=bitmap, name=name, tile=tile, stretch=stretch, auto_name=auto_name)
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
        direct = DirectPattern.from_preset(preset)
        inst = super(Pattern, cls).__new__(cls)
        inst.__init__(style_name=style_name, style_family=style_family)
        inst._set_style("direct", direct, *direct.get_attrs())
        return inst

    @classmethod
    def from_style(
        cls,
        doc: object,
        style_name: StyleParaKind | str = StyleParaKind.STANDARD,
        style_family: str = "ParagraphStyles",
    ) -> Pattern:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Documnet Object.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to. Deftult is Default Paragraph Style.
            style_family (str, optional): Style family. Defatult ``ParagraphStyles``.

        Returns:
            Alignment: ``Alignment`` instance from document properties.
        """
        inst = super(Pattern, cls).__new__(cls)
        inst.__init__(style_name=style_name, style_family=style_family)
        direct = DirectPattern.from_obj(inst.get_style_props(doc))
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
    def prop_inner(self) -> DirectPattern:
        """Gets Inner Pattern instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(DirectPattern, self._get_style_inst("direct"))
        return self._direct_inner
