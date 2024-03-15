from __future__ import annotations
from typing import cast, TYPE_CHECKING, Generic, TypeVar
import uno

from ooodev.mock import mock_g
from ooodev.draw.shapes.draw_shape import DrawShape
from ooodev.draw.shapes.const import KNOWN_SHAPES
from ooodev.loader import lo as mLo
from ooodev.loader.inst.lo_inst import LoInst
from ooodev.draw.shapes.shape_class_factory import ShapeClassFactory
from ooodev.office.partial.office_document_prop_partial import OfficeDocumentPropPartial


_T = TypeVar("_T", bound="ComponentT")


if TYPE_CHECKING:
    from com.sun.star.drawing import XShape
    from ooodev.proto.component_proto import ComponentT
    from ooodev.draw.shapes import ShapeBase


class ShapeFactoryPartial(Generic[_T], OfficeDocumentPropPartial):
    """Partial class for ShapeFactory."""

    def __init__(self, owner: ComponentT, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            owner (_T): Owner object.
            lo_inst (LoInst | None, optional): Lo Instance. Defaults to None.
        """
        self.__owner = owner
        if not isinstance(owner, OfficeDocumentPropPartial):
            raise ValueError("owner must be an instance of OfficeDocumentPropPartial")
        OfficeDocumentPropPartial.__init__(self, owner.office_doc)
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
        # pylint: disable=import-outside-toplevel
        # pylint: disable=redefined-outer-name
        if not hasattr(shape, "ShapeType"):
            raise ValueError("shape has no ShapeType property")
        base_class = None
        shape_type = cast(str, getattr(shape, "ShapeType"))
        if shape_type == "com.sun.star.drawing.ClosedBezierShape":
            from ooodev.draw.shapes.closed_bezier_shape import ClosedBezierShape

            base_class = ClosedBezierShape
        elif shape_type == "com.sun.star.drawing.ConnectorShape":
            from ooodev.draw.shapes.connector_shape import ConnectorShape

            base_class = ConnectorShape
        elif shape_type == "com.sun.star.drawing.EllipseShape":
            from ooodev.draw.shapes.ellipse_shape import EllipseShape

            base_class = EllipseShape
        elif shape_type == "com.sun.star.drawing.GraphicObjectShape":
            from ooodev.draw.shapes.graphic_object_shape import GraphicObjectShape

            base_class = GraphicObjectShape
        elif shape_type == "com.sun.star.drawing.LineShape":
            from ooodev.draw.shapes.line_shape import LineShape

            base_class = LineShape
        elif shape_type == "com.sun.star.drawing.OLE2Shape":
            from ooodev.draw.shapes.ole2_shape import OLE2Shape

            base_class = OLE2Shape
        elif shape_type == "com.sun.star.drawing.OpenBezierShape":
            from ooodev.draw.shapes.open_bezier_shape import OpenBezierShape

            base_class = OpenBezierShape
        elif shape_type == "com.sun.star.drawing.PolyLineShape":
            from ooodev.draw.shapes.poly_line_shape import PolyLineShape

            base_class = PolyLineShape
        elif shape_type == "com.sun.star.drawing.PolyPolygonShape":
            from ooodev.draw.shapes.poly_polygon_shape import PolyPolygonShape

            base_class = PolyPolygonShape
        elif shape_type == "com.sun.star.drawing.RectangleShape":
            from ooodev.draw.shapes.rectangle_shape import RectangleShape

            base_class = RectangleShape
        elif shape_type == "com.sun.star.drawing.TextShape":
            from ooodev.draw.shapes.text_shape import TextShape

            base_class = TextShape
        elif shape_type == "FrameShape":
            from ooodev.write.write_text_frame import WriteTextFrame

            base_class = WriteTextFrame

        if base_class is None:
            class_factory = ShapeClassFactory(owner=self.__owner, component=shape, lo_inst=self.__lo_inst)
        else:
            class_factory = ShapeClassFactory(
                owner=self.__owner, component=shape, lo_inst=self.__lo_inst, base_class=base_class
            )
        return class_factory.get_class()
        # return DrawShape(owner=self.__owner, component=shape, lo_inst=self.__lo_inst)

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
    # pylint: disable=unused-import
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
