from __future__ import annotations
from typing import cast, TYPE_CHECKING, Generic, TypeVar
import uno

from ooodev.mock import mock_g
from ooodev.draw.shapes.draw_shape import DrawShape
from ooodev.draw.shapes.const import KNOWN_SHAPES
from ooodev.loader import lo as mLo
from ooodev.loader.inst.lo_inst import LoInst
from ooodev.proto.component_proto import ComponentT


_T = TypeVar("_T", bound="ComponentT")

if TYPE_CHECKING:
    from com.sun.star.drawing import XShape
    from ooodev.draw.shapes import ShapeBase


class ShapeFactoryPartial(Generic[_T]):
    """Partial class for ShapeFactory."""

    def __init__(self, owner: _T, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            owner (_T): Owner object.
            lo_inst (LoInst | None, optional): Lo Instance. Defaults to None.
        """
        self.__owner = owner
        self.__lo_inst = mLo.Lo.current_lo if lo_inst is None else lo_inst

    def shape_factory(self, shape: XShape) -> ShapeBase[_T]:
        """
        Gets a ShapeBase object from a XShape object.

        Args:
            shape (XShape): UNO object of a shape.

        Raises:
            ValueError: If shape does not have ShapeType property.

        Returns:
            ShapeBase[_T]: Object the inherits from ShapeBase such as ``RectangleShape`` or ``EllipseShape``.
            If ``ShapeType`` of shape is not a match then ``DrawShape[_T]`` is returned.
        """
        if not hasattr(shape, "ShapeType"):
            raise ValueError("shape has no ShapeType property")
        shape_type = cast(str, getattr(shape, "ShapeType"))
        if shape_type == "com.sun.star.drawing.ClosedBezierShape":
            from ooodev.draw.shapes.closed_bezier_shape import ClosedBezierShape

            return ClosedBezierShape(owner=self.__owner, component=shape, lo_inst=self.__lo_inst)
        if shape_type == "com.sun.star.drawing.ConnectorShape":
            from ooodev.draw.shapes.connector_shape import ConnectorShape

            return ConnectorShape(owner=self.__owner, component=shape, lo_inst=self.__lo_inst)
        if shape_type == "com.sun.star.drawing.EllipseShape":
            from ooodev.draw.shapes.ellipse_shape import EllipseShape

            return EllipseShape(owner=self.__owner, component=shape, lo_inst=self.__lo_inst)
        if shape_type == "com.sun.star.drawing.GraphicObjectShape":
            from ooodev.draw.shapes.graphic_object_shape import GraphicObjectShape

            return GraphicObjectShape(owner=self.__owner, component=shape, lo_inst=self.__lo_inst)
        if shape_type == "com.sun.star.drawing.LineShape":
            from ooodev.draw.shapes.line_shape import LineShape

            return LineShape(owner=self.__owner, component=shape, lo_inst=self.__lo_inst)
        if shape_type == "com.sun.star.drawing.OLE2Shape":
            from ooodev.draw.shapes.ole2_shape import OLE2Shape

            return OLE2Shape(owner=self.__owner, component=shape, lo_inst=self.__lo_inst)
        if shape_type == "com.sun.star.drawing.OpenBezierShape":
            from ooodev.draw.shapes.open_bezier_shape import OpenBezierShape

            return OpenBezierShape(owner=self.__owner, component=shape, lo_inst=self.__lo_inst)
        if shape_type == "com.sun.star.drawing.PolyLineShape":
            from ooodev.draw.shapes.poly_line_shape import PolyLineShape

            return PolyLineShape(owner=self.__owner, component=shape, lo_inst=self.__lo_inst)
        if shape_type == "com.sun.star.drawing.PolyPolygonShape":
            from ooodev.draw.shapes.poly_polygon_shape import PolyPolygonShape

            return PolyPolygonShape(owner=self.__owner, component=shape, lo_inst=self.__lo_inst)
        if shape_type == "com.sun.star.drawing.RectangleShape":
            from ooodev.draw.shapes.rectangle_shape import RectangleShape

            return RectangleShape(owner=self.__owner, component=shape, lo_inst=self.__lo_inst)
        if shape_type == "com.sun.star.drawing.TextShape":
            from ooodev.draw.shapes.text_shape import TextShape

            return TextShape(owner=self.__owner, component=shape, lo_inst=self.__lo_inst)
        if shape_type == "FrameShape":
            from ooodev.write.write_text_frame import WriteTextFrame

            return WriteTextFrame(owner=self.__owner, component=shape, lo_inst=self.__lo_inst)  # type: ignore

        return DrawShape(owner=self.__owner, component=shape, lo_inst=self.__lo_inst)

    def is_know_shape(self, shape: XShape) -> bool:
        """
        Checks if a shape is known by this factory.

        Args:
            shape (XShape): UNO object of a shape.

        Returns:
            bool: ``True`` if shape is known; Otherwise, ``False``.
        """
        shape_type = shape.getShapeType()
        return shape_type in KNOWN_SHAPES


if mock_g.FULL_IMPORT:
    from ooodev.draw.shapes.closed_bezier_shape import ClosedBezierShape
    from ooodev.draw.shapes.connector_shape import ConnectorShape
    from ooodev.draw.shapes.ellipse_shape import EllipseShape
    from ooodev.draw.shapes.graphic_object_shape import GraphicObjectShape
    from ooodev.draw.shapes.line_shape import LineShape
    from ooodev.draw.shapes.ole2_shape import OLE2Shape
    from ooodev.draw.shapes.open_bezier_shape import OpenBezierShape
    from ooodev.draw.shapes.poly_line_shape import PolyLineShape
    from ooodev.draw.shapes.poly_polygon_shape import PolyPolygonShape
    from ooodev.draw.shapes.rectangle_shape import RectangleShape
    from ooodev.draw.shapes.text_shape import TextShape
    from ooodev.write.write_text_frame import WriteTextFrame
