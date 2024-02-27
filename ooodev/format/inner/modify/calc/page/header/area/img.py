# region Imports
from __future__ import annotations
from typing import Tuple, cast
import uno
from com.sun.star.awt import XBitmap
from ooo.dyn.drawing.rectangle_point import RectanglePoint

from ooodev.format.calc.style.page.kind.calc_style_page_kind import CalcStylePageKind
from ooodev.format.inner.common.format_types.offset_column import OffsetColumn
from ooodev.format.inner.common.format_types.offset_row import OffsetRow
from ooodev.format.inner.common.format_types.size_percent import SizePercent
from ooodev.format.inner.common.props.img_para_area_props import ImgParaAreaProps
from ooodev.format.inner.direct.write.fill.area.img import ImgStyleKind
from ooodev.format.inner.direct.write.table.background.img import Img as TblImg
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.modify.calc.cell_style_base_multi import CellStyleBaseMulti
from ooodev.format.inner.preset.preset_image import PresetImageKind
from ooodev.utils.data_type.offset import Offset
from ooodev.utils.data_type.size_mm import SizeMM


# endregion Imports


class InnerImg(TblImg):
    """
    Class for Style background image.

    .. seealso::

        - :ref:`help_calc_format_modify_page_header_background`

    .. versionadded:: 0.9.0
    """

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.style.PageProperties", "com.sun.star.style.PageStyle")
        return self._supported_services_values

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.STYLE | FormatKind.PAGE | FormatKind.IMAGE | FormatKind.HEADER
        return self._format_kind_prop

    @property
    def _props(self) -> ImgParaAreaProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = ImgParaAreaProps(
                back_color="HeaderBackColor",
                back_graphic="HeaderBackGraphic",
                graphic_loc="HeaderBackGraphicLocation",
                transparent="HeaderBackTransparent",
            )
        return self._props_internal_attributes


class Img(CellStyleBaseMulti):
    """
    Page Style Image

    .. seealso::

        - :ref:`help_calc_format_modify_page_header_background`

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
        style_name: CalcStylePageKind | str = CalcStylePageKind.DEFAULT,
        style_family: str = "PageStyles",
    ) -> None:
        """
        Constructor

        Args:
            bitmap (XBitmap, optional): Bitmap instance. If ``name`` is not already in the Bitmap Table
                then this property is required.
            name (str, optional): Specifies the name of the image. This is also the name that is used to store
                bitmap in LibreOffice Bitmap Table.
            mode (ImgStyleKind, optional): Specifies the image style, tiled, stretched etc.
                Default ``ImgStyleKind.TILED``.
            size (SizePercent, SizeMM, optional): Size in percent (``0 - 100``) or size in ``mm`` units.
            position (RectanglePoint): Tiling position of Image.
            pos_offset (Offset, optional): Tiling position offset.
            tile_offset (OffsetColumn, OffsetRow, optional): The tiling offset.
            auto_name (bool, optional): Specifies if ``name`` is ensured to be unique. Defaults to ``False``.
            style_name (CalcStylePageKind, str, optional): Specifies the Page Style that instance applies to.
                Default is Default Page Style.
            style_family (str, optional): Style family. Default ``PageStyles``.

        Returns:
            None:

        See Also:
            - :ref:`help_calc_format_modify_page_header_background`
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
            _cattribs=self._get_inner_cattribs(),
        )
        super().__init__()
        self._style_name = str(style_name)
        self._style_family_name = style_family
        self._set_style("direct", direct)

    # region internal methods
    def _get_inner_cattribs(self) -> dict:
        props = ImgParaAreaProps(
            back_color="HeaderBackColor",
            back_graphic="HeaderBackGraphic",
            graphic_loc="HeaderBackGraphicLocation",
            transparent="HeaderBackTransparent",
        )
        return {
            "_props_internal_attributes": props,
            "_supported_services_values": self._supported_services(),
            "_format_kind_prop": self.prop_format_kind,
        }

    # endregion internal methods

    # region Static Methods

    @classmethod
    def from_preset(
        cls,
        preset: PresetImageKind,
        style_name: CalcStylePageKind | str = CalcStylePageKind.DEFAULT,
        style_family: str = "PageStyles",
    ) -> Img:
        """
        Gets an instance from a preset.

        Args:
            preset (PresetImageKind): Preset.
            style_name (CalcStylePageKind, str, optional): Specifies the Paragraph Style that instance applies to.
                Default is Default Paragraph Style.
            style_family (str, optional): Style family. Default ``PageStyles``.

        Returns:
            Img: ``Img`` instance from preset.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerImg.from_preset(preset=preset, _cattribs=inst._get_inner_cattribs())
        inst._set_style("direct", direct)
        return inst

    # endregion Static Methods

    # region Properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.DOC | FormatKind.STYLE | FormatKind.HEADER
        return self._format_kind_prop

    @property
    def prop_style_name(self) -> str:
        """Gets/Sets property Style Name"""
        return self._style_name

    @prop_style_name.setter
    def prop_style_name(self, value: str | CalcStylePageKind):
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
