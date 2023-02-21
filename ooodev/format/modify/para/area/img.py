from __future__ import annotations
from typing import cast
import uno
from ....writer.style.para.kind import StyleParaKind as StyleParaKind
from .....utils.data_type.offset import Offset as Offset
from ...para.para_style_base_multi import ParaStyleBaseMulti
from ....preset.preset_image import PresetImageKind as PresetImageKind
from ....direct.fill.area.img import (
    Img as DirectImg,
    SizeMM as SizeMM,
    SizePercent as SizePercent,
    OffsetColumn as OffsetColumn,
    OffsetRow as OffsetRow,
    ImgStyleKind as ImgStyleKind,
)

from com.sun.star.awt import XBitmap

from ooo.dyn.drawing.rectangle_point import RectanglePoint as RectanglePoint

# Writer Paragraph Style does not need to set paragraph properties, Setting fill properties only
# works whereas setting paragraph properties on style messes up the graphic in this case.


class Img(ParaStyleBaseMulti):
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
        style_name: StyleParaKind | str = StyleParaKind.STANDARD,
        style_family: str = "ParagraphStyles",
    ) -> None:
        """
        Constructor

        Args:
            color (Color, optional): FillColor Color
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to. Deftult is Default Paragraph Style.

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
        self._style_name = str(style_name)
        self._style_family_name = style_family
        self._set_style("direct", direct, *direct.get_attrs())

    @classmethod
    def from_preset(
        cls,
        preset: PresetImageKind,
        style_name: StyleParaKind | str = StyleParaKind.STANDARD,
        style_family: str = "ParagraphStyles",
    ) -> Img:
        """
        Gets an instance from a preset

        Args:
            preset (PresetImageKind): Preset

        Returns:
            Img: Instance from preset.
        """
        direct = DirectImg.from_preset(preset)
        inst = cls(style_name=style_name, style_family=style_family)
        inst._set_style("direct", direct, *direct.get_attrs())

        return inst

    @classmethod
    def from_style(
        cls,
        doc: object,
        style_name: StyleParaKind | str = StyleParaKind.STANDARD,
        style_family: str = "ParagraphStyles",
    ) -> Img:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Documnet Object.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to. Deftult is Default Paragraph Style.
            style_family (str, optional): Style family. Defatult ``ParagraphStyles``.

        Returns:
            Img: ``Img`` instance from document properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = DirectImg.from_obj(inst.get_style_props(doc))
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
    def prop_inner(self) -> DirectImg:
        """Gets Inner Image instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(DirectImg, self._get_style_inst("direct"))
        return self._direct_inner
