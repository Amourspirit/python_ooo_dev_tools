from __future__ import annotations
from typing import Any, cast, overload, TYPE_CHECKING
import uno
from com.sun.star.drawing import XShape
from ooo.dyn.awt.size import Size as UnoSize

from ooodev.exceptions import ex as mEx
from ooodev.format.inner.direct.chart2.position_size.size import Size as ShapeSize
from ooodev.format.inner.direct.draw.shape.position_size.position import Position
from ooodev.loader import lo as mLo
from ooodev.utils import props as mProps
from ooodev.utils.data_type.point import Point as OooDevPoint
from ooodev.utils.data_type.size import Size as OooDevSize
from ooodev.utils.kind.shape_base_point_kind import ShapeBasePointKind

if TYPE_CHECKING:
    from ooodev.units.unit_obj import UnitT


def calculate_x_and_y_from_point_kind(
    x: int, y: int, orig_size: OooDevSize, new_size: OooDevSize, point_kind: ShapeBasePointKind
) -> OooDevPoint:
    """
    Calculates the x and y coordinates from the point kind for a shapes size.

    Args:
        x (int): x in ``1/100th mm``.
        y (int): y in ``1/100th mm``.
        orig_size (Size): Shape original size in ``1/100th mm``.
        new_size (Size): Shape new size in ``1/100th mm``.
        point_kind (ShapeBasePointKind): Point Kind

    Raises:
        ValueError: If Shape size is <= 0.
        ValueError: Unknown point_kind.

    Returns:
        Size: Size in ``1/100th mm``.
    """
    if orig_size.width <= 0:
        raise ValueError(f"orig_size.width must be > 0, not {orig_size.width}")
    if orig_size.height <= 0:
        raise ValueError(f"orig_size.height must be > 0, not {orig_size.height}")

    if new_size.width <= 0:
        raise ValueError(f"new_size.width must be > 0, not {new_size.width}")
    if new_size.height <= 0:
        raise ValueError(f"new_size.height must be > 0, not {new_size.height}")

    if point_kind == ShapeBasePointKind.TOP_LEFT:
        pass
    elif point_kind == ShapeBasePointKind.TOP_CENTER:
        # Shape left will move half the difference -/+
        diff_width = orig_size.width - new_size.width
        if diff_width != 0:
            x += round(diff_width / 2)
    elif point_kind == ShapeBasePointKind.TOP_RIGHT:
        # Shape left will move full difference -/+
        diff_width = orig_size.width - new_size.width
        if diff_width != 0:
            x += diff_width
    elif point_kind == ShapeBasePointKind.CENTER_LEFT:
        # Shape top will move half the difference -/+
        diff_height = orig_size.height - new_size.height
        if diff_height != 0:
            y += round(diff_height / 2)
    elif point_kind == ShapeBasePointKind.CENTER:
        diff_width = orig_size.width - new_size.width
        if diff_width != 0:
            x += round(diff_width / 2)
        diff_height = orig_size.height - new_size.height
        if diff_height != 0:
            y += round(diff_height / 2)
    elif point_kind == ShapeBasePointKind.CENTER_RIGHT:
        diff_width = orig_size.width - new_size.width
        if diff_width != 0:
            x += diff_width
        diff_height = orig_size.height - new_size.height
        if diff_height != 0:
            y += round(diff_height / 2)
    elif point_kind == ShapeBasePointKind.BOTTOM_LEFT:
        diff_height = orig_size.height - new_size.height
        if diff_height != 0:
            y += diff_height
    elif point_kind == ShapeBasePointKind.BOTTOM_CENTER:
        diff_width = orig_size.width - new_size.width
        if diff_width != 0:
            x += round(diff_width / 2)
        diff_height = orig_size.height - new_size.height
        if diff_height != 0:
            y += diff_height
    elif point_kind == ShapeBasePointKind.BOTTOM_RIGHT:
        diff_width = orig_size.width - new_size.width
        if diff_width != 0:
            x += diff_width
        diff_height = orig_size.height - new_size.height
        if diff_height != 0:
            y += diff_height
    else:
        raise ValueError(f"Unknown point_kind: {point_kind}")
    return OooDevPoint(x=x, y=y)


class Size(ShapeSize):
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
        width: float | UnitT,
        height: float | UnitT,
        base_point: ShapeBasePointKind = ShapeBasePointKind.TOP_LEFT,
    ) -> None:
        """
        Constructor

        Args:
            width (float | UnitT): Specifies the width of the shape (in ``mm`` units) or :ref:`proto_unit_obj`.
            height (float | UnitT): Specifies the height of the shape (in ``mm`` units) or :ref:`proto_unit_obj`.
            base_point (ShapeBasePointKind): Specifies the base point of the shape used to calculate the X and Y coordinates. Default is ``TOP_LEFT``.

        Returns:
            None:
        """
        super().__init__(width=width, height=height)
        self._base_point = base_point

        # draw_page.BorderBottom

    # region Overridden Methods
    def apply(self, obj: Any, **kwargs) -> None:
        """
        Applies properties to ``obj``

        Args:
            obj (Any): UNO object.

        Returns:
            None:
        """
        name = self._get_property_name()
        if not name:
            return
        shape = mLo.Lo.qi(XShape, obj)
        if shape is None:
            return
        new_size = OooDevSize(width=self._width, height=self._height)
        if self._base_point != ShapeBasePointKind.TOP_LEFT:
            sz = shape.getSize()
            orig_size = OooDevSize(width=sz.Width, height=sz.Height)
            pos = Position.from_obj(obj=obj)
            pt = calculate_x_and_y_from_point_kind(
                x=pos._pos_x, y=pos._pos_y, orig_size=orig_size, new_size=new_size, point_kind=self._base_point
            )
            if pt.x != pos._pos_x or pt.y != pos._pos_y:
                # there is a difference
                pos._pos_x = pt.x
                pos._pos_y = pt.y
                pos.apply(obj=obj)

        # Draw automatically add the page borders to the position when setting.
        # struct = UnoPoint(X=size.width + self._draw_page.BorderLeft, Y=size.height + self._draw_page.BorderTop)
        struct = UnoSize(Width=self._width, Height=self._height)
        props = kwargs.pop("override_dv", {})
        props.update({name: struct})
        super().apply(obj=obj, override_dv=props, update_dv=False)

    def copy(self, **kwargs) -> Size:
        """
        Copy the current instance.

        Returns:
            Size: The copied instance.
        """
        # pylint: disable=protected-access
        cp = cast(Size, super().copy(**kwargs))
        cp._base_point = self._base_point
        return cp

    # endregion Overridden Methods

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls, obj: Any) -> Size:
        """
        Gets size from ``obj``

        Args:
            obj (Any): UNO Shape object.

        Returns:
            Size: New instance.
        """
        ...

    @overload
    @classmethod
    def from_obj(cls, obj: Any, **kwargs) -> Size:
        """
        Gets size from ``obj``

        Args:
            obj (Any): UNO Shape object.
            **kwargs: Additional arguments.

        Returns:
            Size: New instance.
        """
        ...

    @classmethod
    def from_obj(cls, obj: Any, **kwargs) -> Size:
        """
        Gets size from ``obj``

        Args:
            obj (Any): UNO Shape object.

        Returns:
            Size: New instance.
        """
        inst = cls(width=0, height=0, **kwargs)
        name = inst._get_property_name()
        if not name:
            raise mEx.NotSupportedError("Object is not supported for conversion to Size")

        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError("Object is not supported for conversion to Size")

        sz = cast(UnoSize, mProps.Props.get(obj, name, None))
        if sz is not None:
            inst._width = sz.Width
            inst._height = sz.Height
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
