# region Imports
from __future__ import annotations
from typing import cast
from com.sun.star.awt import XBitmap
from ooo.dyn.drawing.rectangle_point import RectanglePoint

from ooodev.format.inner.common.format_types.offset_column import OffsetColumn
from ooodev.format.inner.common.format_types.offset_row import OffsetRow
from ooodev.format.inner.common.format_types.size_percent import SizePercent
from ooodev.format.inner.direct.write.fill.area.img import Img as InnerImg
from ooodev.format.inner.direct.write.fill.area.img import ImgStyleKind
from ooodev.format.inner.modify.write.frame.frame_style_base_multi import FrameStyleBaseMulti
from ooodev.format.inner.preset.preset_image import PresetImageKind
from ooodev.format.writer.style.frame.style_frame_kind import StyleFrameKind
from ooodev.utils.data_type.offset import Offset
from ooodev.utils.data_type.size_mm import SizeMM

# endregion Imports


class Img(FrameStyleBaseMulti):
    """
    Frame Style Area Image.

    .. versionadded:: 0.9.0
    """

    # region Init
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
        style_name: StyleFrameKind | str = StyleFrameKind.FRAME,
        style_family: str = "FrameStyles",
    ) -> None:
        """
        Constructor

        Args:
            bitmap (XBitmap, optional): Bitmap instance. If ``name`` is not already in the Bitmap Table then this property is required.
            name (str, optional): Specifies the name of the image. This is also the name that is used to store bitmap in LibreOffice Bitmap Table.
            mode (ImgStyleKind, optional): Specifies the image style, tiled, stretched etc. Default ``ImgStyleKind.TILED``.
            size (SizePercent, SizeMM, optional): Size in percent (``0 - 100``) or size in ``mm`` units.
            position (RectanglePoint): Tiling position of Image.
            pos_offset (Offset, optional): Tiling position offset.
            tile_offset (OffsetColumn, OffsetRow, optional): The tiling offset.
            auto_name (bool, optional): Specifies if ``name`` is ensured to be unique. Defaults to ``False``.
            style_name (StyleFrameKind, str, optional): Specifies the Frame Style that instance applies to. Default is Default Frame Style.
            style_family (str, optional): Style family. Default ``FrameStyles``.

        Returns:
            None:
        """
        # pylint: disable=protected-access
        # pylint: disable=unexpected-keyword-arg
        direct = InnerImg(
            bitmap=bitmap,
            name=name,
            mode=mode,
            size=size,
            position=position,
            pos_offset=pos_offset,
            tile_offset=tile_offset,
            auto_name=auto_name,
            _cattribs=self._get_inner_cattribs(),
        )
        direct._prop_parent = self
        super().__init__()
        self._style_name = str(style_name)
        self._style_family_name = style_family
        self._set_style("direct", direct, *direct.get_attrs())

    # endregion Init

    # region Internal Methods
    def _get_inner_cattribs(self) -> dict:
        return {"_supported_services_values": self._supported_services(), "_format_kind_prop": self.prop_format_kind}

    # endregion Internal Methods

    # region Static Methods
    @classmethod
    def from_style(
        cls,
        doc: object,
        style_name: StyleFrameKind | str = StyleFrameKind.FRAME,
        style_family: str = "FrameStyles",
    ) -> Img:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Document Object.
            style_name (StyleFrameKind, str, optional): Specifies the Frame Style that instance applies to. Default is Default Frame Style.
            style_family (str, optional): Style family. Default ``FrameStyles``.

        Returns:
            Img: ``Img`` instance from style properties.
        """
        # pylint: disable=protected-access
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerImg.from_obj(obj=inst.get_style_props(doc), _cattribs=inst._get_inner_cattribs())
        direct._prop_parent = inst
        inst._set_style("direct", direct, *direct.get_attrs())
        return inst

    @classmethod
    def from_preset(
        cls,
        preset: PresetImageKind,
        style_name: StyleFrameKind | str = StyleFrameKind.FRAME,
        style_family: str = "FrameStyles",
    ) -> Img:
        """
        Gets an instance from a preset.

        Args:
            preset (PresetImageKind): Preset.
            style_name (StyleFrameKind, str, optional): Specifies the Frame Style that instance applies to. Default is Default Frame Style.
            style_family (str, optional): Style family. Default ``FrameStyles``.

        Returns:
            Img: ``Img`` instance from preset.
        """
        # pylint: disable=protected-access
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerImg.from_preset(preset=preset, _cattribs=inst._get_inner_cattribs())
        direct._prop_parent = inst
        inst._set_style("direct", direct, *direct.get_attrs())
        return inst

    # endregion Static Methods

    # region Properties
    @property
    def prop_style_name(self) -> str:
        """Gets/Sets property Style Name"""
        return self._style_name

    @prop_style_name.setter
    def prop_style_name(self, value: str | StyleFrameKind):
        self._style_name = str(value)

    @property
    def prop_inner(self) -> InnerImg:
        """Gets Inner Image instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(InnerImg, self._get_style_inst("direct"))
        return self._direct_inner

    # endregion Properties
