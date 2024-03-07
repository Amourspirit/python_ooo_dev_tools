# region Imports
from __future__ import annotations
from typing import Any, Tuple, overload

import uno
from com.sun.star.awt import XBitmap
from com.sun.star.graphic import XGraphic

from ooo.dyn.drawing.rectangle_point import RectanglePoint

from ooodev.events.args.key_val_cancel_args import KeyValCancelArgs
from ooodev.loader import lo as mLo
from ooodev.utils.data_type.offset import Offset
from ooodev.format.inner.direct.write.fill.area.img import ImgStyleKind
from ooodev.format.inner.common.format_types.offset_row import OffsetRow
from ooodev.format.inner.common.format_types.offset_column import OffsetColumn
from ooodev.format.inner.common.format_types.size_percent import SizePercent
from ooodev.utils.data_type.size_mm import SizeMM
from ooodev.format.inner.common.props.img_para_area_props import ImgParaAreaProps
from ooodev.meta.disabled_method import DisabledMethod
from ooodev.format.inner.direct.write.para.area.img import Img as ParaImg

# endregion Imports


class Img(ParaImg):
    """
    Class for table background image.

    .. seealso::

        - :ref:`help_writer_format_direct_table_background`

    .. versionadded:: 0.9.0
    """

    from_obj = DisabledMethod()  # type: ignore[assignment] #
    """From object is not supported in this class."""

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

        Returns:
            None:

        Note:
            If ``auto_name`` is ``False`` then a bitmap for a given ``name`` is only required the first call.
            All subsequent call of the same ``name`` will retrieve the bitmap form the LibreOffice Bitmap Table.

        See Also:

            - :ref:`help_writer_format_direct_table_borders`
        """
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

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = (
                "com.sun.star.text.TextTable",
                "com.sun.star.text.CellProperties",
                "com.sun.star.text.TextTableRow",
            )
        return self._supported_services_values

    def _on_multi_child_style_applying(self, source: Any, event_args: KeyValCancelArgs) -> None:
        if event_args.key == "fill_image":
            event_args.cancel = True
        return super()._on_multi_child_style_applying(source, event_args)

    # region apply()

    @overload
    def apply(self, obj: Any) -> None: ...

    @overload
    def apply(self, obj: Any, **kwargs) -> None: ...

    def apply(self, obj: Any, **kwargs) -> None:
        """
        Applies styles to object

        Args:
            obj (object): UNO object.

        Returns:
            None:
        """
        self._clear()
        self._set(self._props.back_color, -1)
        self._set(self._props.transparent, True)
        # will not work if not first converted to XGraphic
        graphic = mLo.Lo.qi(XGraphic, self.prop_inner.prop_bitmap, True)
        self._set(self._props.back_graphic, graphic)
        loc = self._get_graphic_loc(position=None, mode=self.prop_inner.prop_mode)
        self._set(self._props.graphic_loc, loc)
        super().apply(obj, **kwargs)

    # endregion apply()

    @property
    def _props(self) -> ImgParaAreaProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = ImgParaAreaProps(
                back_color="BackColor",
                back_graphic="BackGraphic",
                graphic_loc="BackGraphicLocation",
                transparent="BackTransparent",
            )
        return self._props_internal_attributes
