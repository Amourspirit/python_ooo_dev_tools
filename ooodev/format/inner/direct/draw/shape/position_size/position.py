from __future__ import annotations
import uno
from com.sun.star.lang import XComponent
from ooodev.format.inner.style_base import StyleMulti
from ooodev.utils.kind.shape_base_point_kind import ShapeBasePointKind
from ooodev.utils.data_type.size import Size
from ooodev.units import UnitT, UnitConvert
from ooodev.format.inner.direct.structs.point_struct import PointStruct


def calculate_x_and_y_from_point_kind(x: int, y: int, shape_size: Size, point_kind: ShapeBasePointKind) -> Size:
    if shape_size.width <= 0:
        raise ValueError(f"shape_size.width must be > 0, not {shape_size.width}")
    if shape_size.height <= 0:
        raise ValueError(f"shape_size.height must be > 0, not {shape_size.height}")

    if point_kind == ShapeBasePointKind.TOP_LEFT:
        pass
    elif point_kind == ShapeBasePointKind.TOP_CENTER:
        x -= round(shape_size.width / 2)
    elif point_kind == ShapeBasePointKind.TOP_RIGHT:
        x = x - shape_size.width
    elif point_kind == ShapeBasePointKind.CENTER_LEFT:
        y -= round(shape_size.height / 2)
    elif point_kind == ShapeBasePointKind.CENTER:
        x = x - round(shape_size.width / 2)
        y = y - round(shape_size.height / 2)
    elif point_kind == ShapeBasePointKind.CENTER_RIGHT:
        x = x - shape_size.width
        y = y - round(shape_size.height / 2)
    elif point_kind == ShapeBasePointKind.BOTTOM_LEFT:
        y = y - shape_size.height
    elif point_kind == ShapeBasePointKind.BOTTOM_CENTER:
        x = x - round(shape_size.width / 2)
        y = y - shape_size.height
    elif point_kind == ShapeBasePointKind.BOTTOM_RIGHT:
        x = x - shape_size.width
        y = y - shape_size.height
    else:
        raise ValueError(f"Unknown point_kind: {point_kind}")
    return Size(width=x, height=y)


class Position(StyleMulti):
    """
    Positions a shape.

    .. versionadded:: 0.9.4
    """

    # in draw document the page margins are also included in the position.
    # if page margin is 10mm and shape is positioned at 0,0 in the dialog box then the shape is actually at 10,10 in the position struct.
    # this means that the position class must be aware of the document margins and add them to the position.
    # for a chart the margins are not included in the position struct.

    # TODO: finish the Position class.

    def __init__(
        self,
        draw_doc: XComponent,
        *,
        pos_x: float | UnitT,
        pos_y: float | UnitT,
        base_point: ShapeBasePointKind = ShapeBasePointKind.TOP_LEFT,
    ) -> None:
        """
        Constructor

        Args:
            pos_x (float | UnitT): Specifies the x-coordinate of the position of the shape (in ``mm`` units) or :ref:`proto_unit_obj`.
            pos_y (float | UnitT): Specifies the y-coordinate of the position of the shape (in ``mm`` units) or :ref:`proto_unit_obj`.
            base_point (ShapeBasePointKind): Specifies the base point of the shape used to calculate the X and Y coordinates. Default is ``TOP_LEFT``.
        """
        super().__init__()
        try:
            self._pos_x = pos_x.get_value_mm100()  # type: ignore
        except AttributeError:
            self._pos_x = UnitConvert.convert_mm_mm100(pos_x)  # type: ignore
        self._pos_y = pos_y

    def _get_inner_point_struct(self) -> PointStruct:
        return PointStruct(x=self._pos_x, y=self._pos_y)
