# region Imports
from __future__ import annotations
from typing import Any, Tuple, cast, overload
import uno
from com.sun.star.awt import XBitmap
from com.sun.star.chart2 import XChartDocument
from com.sun.star.lang import XMultiServiceFactory

from ooo.dyn.drawing.rectangle_point import RectanglePoint

from ooodev.format.inner.common.format_types.offset_column import OffsetColumn as OffsetColumn
from ooodev.format.inner.common.format_types.offset_row import OffsetRow as OffsetRow
from ooodev.format.inner.common.format_types.size_percent import SizePercent as SizePercent
from ooodev.format.inner.common.props.area_img_props import AreaImgProps
from ooodev.format.inner.direct.write.fill.area.img import Img as FillImg, ImgStyleKind
from ooodev.format.inner.preset import preset_image as mImage
from ooodev.format.inner.preset.preset_image import PresetImageKind as PresetImageKind
from ooodev.meta.deleted_attrib import DeletedAttrib
from ooodev.loader import lo as mLo
from ooodev.utils.data_type.offset import Offset as Offset
from ooodev.utils.data_type.size_mm import SizeMM as SizeMM

# endregion Imports


class Img(FillImg):
    """
    Class for Chart Area Fill Image.

    .. seealso::

        - :ref:`help_chart2_format_direct_general_area`

    .. versionadded:: 0.9.4
    """

    prop_bitmap = DeletedAttrib()  # type: ignore

    def __init__(
        self,
        chart_doc: XChartDocument,
        *,
        bitmap: XBitmap | None = None,
        name: str = "",
        mode: ImgStyleKind = ImgStyleKind.TILED,
        size: SizePercent | SizeMM | None = None,
        position: RectanglePoint | None = None,
        pos_offset: Offset | None = None,
        tile_offset: OffsetColumn | OffsetRow | None = None,
        auto_name: bool = False,
    ) -> None:
        """
        Constructor

        Args:
            chart_doc (XChartDocument): Chart document.
            bitmap (XBitmap, optional): Bitmap instance. If ``name`` is not already in the Bitmap Table
                then this property is required.
            name (str, optional): Specifies the name of the image. This is also the name that is used to store
                bitmap in LibreOffice Bitmap Table.
            mode (~.write.fill.area.img.ImgStyleKind, optional): Specifies the image style, tiled, stretched etc.
                Default ``ImgStyleKind.TILED``.
            size (SizePercent, SizeMM, optional): Size in percent (``0 - 100``) or size in ``mm`` units.
            position (RectanglePoint): Tiling position of Image.
            pos_offset (Offset, optional): Tiling position offset.
            tile_offset (OffsetColumn, OffsetRow, optional): The tiling offset.
            auto_name (bool, optional): Specifies if ``name`` is ensured to be unique. Defaults to ``False``.

        Returns:
            None:

        Note:
            If ``auto_name`` is ``False`` then a bitmap for a given ``name`` is only required the first call.
            All subsequent call of the same ``name`` will retrieve the bitmap form the LibreOffice Bitmap Table.

        See Also:
            - :ref:`help_chart2_format_direct_general_area`
        """
        self._chart_doc = chart_doc
        super().__init__(
            bitmap=bitmap,
            name=name,
            mode=mode,
            size=size,
            position=position,
            pos_offset=pos_offset,
            tile_offset=tile_offset,
            auto_name=auto_name,
        )

    # region overrides
    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = (
                "com.sun.star.chart2.DataPoint",
                "com.sun.star.chart2.DataSeries",
                "com.sun.star.chart2.Legend",
                "com.sun.star.chart2.PageBackground",
                "com.sun.star.chart2.Title",
                "com.sun.star.drawing.FillProperties",
            )
        return self._supported_services_values

    def _container_get_msf(self) -> XMultiServiceFactory | None:
        if self._chart_doc is not None:
            return mLo.Lo.qi(XMultiServiceFactory, self._chart_doc)
        return None

    # region copy()
    @overload
    def copy(self) -> Img: ...

    @overload
    def copy(self, **kwargs) -> Img: ...

    def copy(self, **kwargs) -> Img:
        """Gets a copy of instance as a new instance"""
        cp = super().copy(**kwargs)
        cp._chart_doc = self._chart_doc
        return cp

    # endregion copy()

    # endregion overrides

    # region Static Methods
    # region from_preset()
    @overload
    @classmethod
    def from_preset(cls, chart_doc: XChartDocument, preset: PresetImageKind) -> Img: ...

    @overload
    @classmethod
    def from_preset(cls, chart_doc: XChartDocument, preset: PresetImageKind, **kwargs) -> Img: ...

    @classmethod
    def from_preset(cls, chart_doc: XChartDocument, preset: PresetImageKind, **kwargs) -> Img:
        """
        Gets an instance from a preset.

        Args:
            chart_doc (XChartDocument): Chart document.
            preset (~.preset.preset_image.PresetImageKind): Preset.

        Returns:
            Img: Instance from preset.
        """
        name = str(preset)
        nu = cls(chart_doc=chart_doc, **kwargs)

        nc = nu._container_get_inst()
        bitmap = cast(XBitmap, nu._container_get_value(name, nc))
        if bitmap is None:
            bitmap = mImage.get_prest_bitmap(preset)
        inst = cls(
            chart_doc=chart_doc,
            bitmap=bitmap,
            name=name,
            mode=ImgStyleKind.TILED,
            position=RectanglePoint.MIDDLE_MIDDLE,
            pos_offset=Offset(0, 0),
            tile_offset=OffsetRow(0),
            auto_name=False,
            **kwargs,
        )
        # set size
        point = preset._get_point()
        inst._set(inst._props.size_x, point.x)
        inst._set(inst._props.size_y, point.y)
        return inst

    # endregion from_preset()
    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls, chart_doc: XChartDocument, obj: Any) -> Img: ...

    @overload
    @classmethod
    def from_obj(cls, chart_doc: XChartDocument, obj: Any, **kwargs) -> Img: ...

    @classmethod
    def from_obj(cls, chart_doc: XChartDocument, obj: Any, **kwargs) -> Img:
        """
        Gets instance from object

        Args:
            chart_doc (XChartDocument): Chart document.
            obj (object): UNO object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            Img: ``Img`` instance that represents ``obj`` fill image.
        """
        return super().from_obj(obj=obj, chart_doc=chart_doc, **kwargs)

    # endregion from_obj()
    # endregion Static Methods

    # region Properties
    @property
    def _props(self) -> AreaImgProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = AreaImgProps(
                name="FillBitmapName",
                style="FillStyle",
                mode="FillBitmapMode",
                point="FillBitmapRectanglePoint",
                bitmap="",
                offset_x="FillBitmapOffsetX",
                offset_y="FillBitmapOffsetY",
                pos_x="FillBitmapPositionOffsetX",
                pos_y="FillBitmapPositionOffsetY",
                size_x="FillBitmapSizeX",
                size_y="FillBitmapSizeY",
            )
        return self._props_internal_attributes

    # endregion Properties
