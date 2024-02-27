from __future__ import annotations
from typing import Any, cast, overload, TYPE_CHECKING
import uno
from com.sun.star.drawing import XShape
from com.sun.star.awt import Rectangle

from ooo.dyn.awt.point import Point as UnoPoint
from ooodev.exceptions import ex as mEx
from ooodev.format.inner.direct.chart2.position_size.position import Position as ShapePosition
from ooodev.loader import lo as mLo
from ooodev.utils import info as mInfo
from ooodev.utils import props as mProps
from ooodev.utils.data_type.point import Point as OooDevPoint
from ooodev.utils.data_type.size import Size
from ooodev.utils.kind.shape_base_point_kind import ShapeBasePointKind


if TYPE_CHECKING:
    from ooodev.units.unit_obj import UnitT
    from com.sun.star.drawing import DrawPage
else:
    DrawPage = Any


def calculate_point_from_point_kind(x: int, y: int, shape_size: Size, point_kind: ShapeBasePointKind) -> OooDevPoint:
    """
    Calculates the x and y coordinates from the point kind for a shapes size.

    Args:
        x (int): x in ``1/100th mm``.
        y (int): y in ``1/100th mm``.
        shape_size (Size): Shape size in ``1/100th mm``.
        point_kind (ShapeBasePointKind): Point Kind

    Raises:
        ValueError: If Shape size is <= 0.
        ValueError: Unknown point_kind.

    Returns:
        Size: Size in ``1/100th mm``.
    """
    if shape_size.width <= 0:
        raise ValueError(f"shape_size.width must be > 0, not {shape_size.width}")
    if shape_size.height <= 0:
        raise ValueError(f"shape_size.height must be > 0, not {shape_size.height}")

    if point_kind == ShapeBasePointKind.TOP_LEFT:
        pass
    elif point_kind == ShapeBasePointKind.TOP_CENTER:
        x += round(shape_size.width / 2)
    elif point_kind == ShapeBasePointKind.TOP_RIGHT:
        x = x + shape_size.width
    elif point_kind == ShapeBasePointKind.CENTER_LEFT:
        y += round(shape_size.height / 2)
    elif point_kind == ShapeBasePointKind.CENTER:
        x += round(shape_size.width / 2)
        y += round(shape_size.height / 2)
    elif point_kind == ShapeBasePointKind.CENTER_RIGHT:
        x = x + shape_size.width
        y += round(shape_size.height / 2)
    elif point_kind == ShapeBasePointKind.BOTTOM_LEFT:
        y = y + shape_size.height
    elif point_kind == ShapeBasePointKind.BOTTOM_CENTER:
        x += round(shape_size.width / 2)
        y = y + shape_size.height
    elif point_kind == ShapeBasePointKind.BOTTOM_RIGHT:
        x = x + shape_size.width
        y = y + shape_size.height
    else:
        raise ValueError(f"Unknown point_kind: {point_kind}")
    return OooDevPoint(x=x, y=y)


class Position(ShapePosition):
    """
    Positions a shape.

    .. versionadded:: 0.17.3
    """

    # in draw document the page margins are also included in the position.
    # if page margin is 10mm and shape is positioned at 0,0 in the dialog box then the shape is actually at 10,10 in the position struct.
    # this means that the position class must be aware of the document margins and add them to the position.
    # for a chart the margins are not included in the position struct.

    def __init__(
        self,
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

        Returns:
            None:

        Note:
            ``pos_x`` and ``pos_y`` are the coordinates of the shape inside the draw page borders.
            This is the same behavior as the dialog box.
            If the draw page has a border of 10mm and the shape is positioned at 0mm,0mm in the dialog box then the shape
            is actually at 10mm,10mm relative to the draw page document.
        """
        super().__init__(pos_x=pos_x, pos_y=pos_y)
        self._base_point = base_point

        # draw_page.BorderBottom

    # region Overridden Methods
    def apply(self, obj: Any, **kwargs) -> None:
        """
        Applies position properties to ``obj``

        Args:
            obj (Any): UNO object.

        Returns:
            None:
        """
        shape = mLo.Lo.qi(XShape, obj)
        if shape is None:
            return
        offset_x = 0
        offset_y = 0
        parent = cast("DrawPage", mProps.Props.get(shape, "Parent", None))
        if parent is not None and mInfo.Info.support_service(parent, "com.sun.star.drawing.DrawPage"):
            offset_x = parent.BorderLeft
            offset_y = parent.BorderTop
        if self._base_point != ShapeBasePointKind.TOP_LEFT:
            sz = shape.getSize()
            pt = calculate_point_from_point_kind(
                x=self._pos_x, y=self._pos_y, shape_size=Size(sz.Width, sz.Height), point_kind=self._base_point
            )
        else:
            pt = OooDevPoint(x=self._pos_x, y=self._pos_y)

        name = self._get_property_name()
        if not name:
            return

        # Draw automatically add the page borders to the position when setting.
        # struct = UnoPoint(X=size.width + self._draw_page.BorderLeft, Y=size.height + self._draw_page.BorderTop)
        struct = UnoPoint(X=pt.x + offset_x, Y=pt.y + offset_y)
        props = kwargs.pop("override_dv", {})
        props.update({name: struct})
        super().apply(obj=obj, override_dv=props, update_dv=False)

    def copy(self, **kwargs) -> Position:
        """
        Copy the current instance.

        Returns:
            Position: The copied instance.
        """
        # pylint: disable=protected-access
        cp = cast(Position, super().copy(**kwargs))
        cp._base_point = self._base_point
        return cp

    # endregion Overridden Methods

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls, obj: Any) -> Position:
        """
        Creates a new instance from ``obj``.

        Args:
            obj (Any): UNO Shape object.

        Returns:
            Position: New instance.
        """
        ...

    @overload
    @classmethod
    def from_obj(cls, obj: Any, **kwargs) -> Position:
        """
        Creates a new instance from ``obj``.

        Args:
            obj (Any): UNO Shape object.
            **kwargs: Additional arguments.

        Returns:
            Position: New instance.
        """
        ...

    @classmethod
    def from_obj(cls, obj: Any, **kwargs) -> Position:
        """
        Creates a new instance from ``obj``.

        Args:
            obj (Any): UNO Shape object.

        Returns:
            Position: New instance.
        """
        # pylint: disable=protected-access
        shape = mLo.Lo.qi(XShape, obj, True)

        inst = cls(pos_x=0, pos_y=0, **kwargs)

        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError("Object is not supported for conversion to Position")

        offset_x = 0
        offset_y = 0
        # When setting the position of a  shape in a draw document the page borders are also included in the position.
        # This is done automatically by the draw application.
        # However when reading the position of a shape the page borders are not subtracted from the position.
        # For this reason the page borders must be subtracted from the position when reading the position of a shape.
        parent = cast("DrawPage", mProps.Props.get(shape, "Parent", None))
        if parent is not None and mInfo.Info.support_service(parent, "com.sun.star.drawing.DrawPage"):
            offset_x = parent.BorderLeft
            offset_y = parent.BorderTop
        # 'com.sun.star.drawing.DrawPage'

        # Position is affected by rotation and may not be what is expected.
        # BoundRect is not affected by rotation and is the same as the dialog box,
        # it seems to be the same for FrameRect
        # name = inst._get_property_name()
        name = "BoundRect"  # BoundRect is read only
        rect = cast(Rectangle, mProps.Props.get(obj, name, None))
        if rect is not None:
            inst._pos_x = rect.X - offset_x
            inst._pos_y = rect.Y - offset_y
        inst.set_update_obj(obj)
        return inst

    # endregion from_obj()

    # region properties
    @property
    def prop_base_point(self) -> ShapeBasePointKind:
        """
        Gets/Sets the base point of the shape used to calculate the X and Y coordinates.

        Returns:
            ShapeBasePointKind: Base point.
        """
        return self._base_point

    @prop_base_point.setter
    def prop_base_point(self, value: ShapeBasePointKind) -> None:
        self._base_point = value

    # endregion properties
