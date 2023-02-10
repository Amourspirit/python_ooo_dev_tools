from __future__ import annotations
import uno
from ....writer.style.para.kind import StyleParaKind as StyleParaKind
from .....utils.data_type.offset import Offset as Offset
from ..fill_style_base_multi import FillStyleBaseMulti
from ....preset.preset_image import PresetImageKind as PresetImageKind
from ....direct.fill.area.img import (
    Img as DirectImg,
    SizeMM as SizeMM,
    SizePercent as SizePercent,
    OffsetColumn as OffsetColumn,
    OffsetRow as OffsetRow,
    ImgStyleKind as ImgStyleKind,
)
from ....draw.style.kind import DrawStyleFamilyKind as DrawStyleFamilyKind
from ....draw.style.lookup import FamilyGraphics

# from ....direct.para.area.img import Img as DirectImg

from com.sun.star.awt import XBitmap

from ooo.dyn.drawing.rectangle_point import RectanglePoint as RectanglePoint


class Img(FillStyleBaseMulti):
    """
    Paragraph Style Fill Coloring

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        bitmap: XBitmap | None = None,
        name: str = "",
        mode: ImgStyleKind = ImgStyleKind.TILED,
        size: SizePercent | SizeMM | None = None,
        position: RectanglePoint | None = None,
        pos_offset: Offset | None = None,
        tile_offset: OffsetColumn | OffsetRow | None = None,
        auto_name: bool = False,
        style_name: str = FamilyGraphics.DEFAULT_DRAWING_STYLE,
        style_family: str | DrawStyleFamilyKind = DrawStyleFamilyKind.GRAPHICS,
    ) -> None:
        """
        Constructor

        Args:
            color (Color, optional): FillColor Color
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to. Deftult is Default Paragraph Style.
            style_family (str, DrawStyleFamilyKind, optional): Family Style.

        Returns:
            None:
        """

        direct = DirectImg(
            bitmap=bitmap,
            name=name,
            mode=mode,
            size=size,
            position=position,
            pos_offset=pos_offset,
            tile_offset=tile_offset,
            auto_name=auto_name,
        )
        super().__init__()
        self._style_family_name = str(style_family)
        self._style_name = str(style_name)
        self._set_style("direct", direct, *direct.get_attrs())

    @classmethod
    def from_preset(
        cls,
        preset: PresetImageKind,
        style_name: str = FamilyGraphics.DEFAULT_DRAWING_STYLE,
        style_family: str | DrawStyleFamilyKind = DrawStyleFamilyKind.GRAPHICS,
    ) -> Img:
        """
        Gets an instance from a preset

        Args:
            preset (PresetImageKind): Preset

        Returns:
            Img: Instance from preset.
        """
        direct = DirectImg.from_preset(preset)
        inst = super(Img, cls).__new__(cls)
        inst.__init__(style_name=style_name, style_family=style_family)
        inst._set_style("direct", direct, *direct.get_attrs())
        return inst

    @property
    def prop_style_name(self) -> str:
        """Gets/Sets property Style Name"""
        return self._style_name

    @prop_style_name.setter
    def prop_style_name(self, value: str | StyleParaKind):
        self._style_name = str(value)
