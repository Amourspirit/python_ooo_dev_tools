# region Import
from __future__ import annotations
from typing import cast
import uno
from com.sun.star.awt import XBitmap

from ooodev.format.writer.style.para.kind import StyleParaKind as StyleParaKind
from ooodev.utils.data_type.offset import Offset as Offset
from ...para.para_style_base_multi import ParaStyleBaseMulti
from ooodev.format.inner.preset.preset_image import PresetImageKind as PresetImageKind
from ooodev.format.inner.direct.write.fill.area.img import SizeMM as SizeMM
from ooodev.format.inner.direct.write.fill.area.img import SizePercent as SizePercent
from ooodev.format.inner.direct.write.fill.area.img import OffsetColumn as OffsetColumn
from ooodev.format.inner.direct.write.fill.area.img import OffsetRow as OffsetRow
from ooodev.format.inner.direct.write.fill.area.img import ImgStyleKind as ImgStyleKind
from ooodev.format.inner.direct.write.fill.area.img import Img as InnerImg
from ooo.dyn.drawing.rectangle_point import RectanglePoint as RectanglePoint

# endregion Import

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
            color (:py:data:`~.utils.color.Color`, optional): FillColor Color
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to.
                Default is Default Paragraph Style.
            style_family (str, optional): Style family. Default ``ParagraphStyles``.

        Returns:
            None:
        """

        direct = InnerImg(
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
        direct = InnerImg.from_preset(preset)
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
            doc (object): UNO Document Object.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to.
                Default is Default Paragraph Style.
            style_family (str, optional): Style family. Default ``ParagraphStyles``.

        Returns:
            Img: ``Img`` instance from document properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerImg.from_obj(inst.get_style_props(doc))
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
    def prop_inner(self) -> InnerImg:
        """Gets/Sets Inner Image instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(InnerImg, self._get_style_inst("direct"))
        return self._direct_inner

    @prop_inner.setter
    def prop_inner(self, value: InnerImg) -> None:
        if not isinstance(value, InnerImg):
            raise TypeError(f'Expected type of InnerImg, got "{type(value).__name__}"')
        self._del_attribs("_direct_inner")
        self._set_style("direct", value, *value.get_attrs())
