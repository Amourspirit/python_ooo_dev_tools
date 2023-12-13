from __future__ import annotations
from typing import Any, List, Tuple, overload, Sequence, TYPE_CHECKING, TypeVar, Generic, Union
import uno


from com.sun.star.drawing import XDrawPage

from ooo.dyn.awt.point import Point
from ooo.dyn.drawing.polygon_flags import PolygonFlags

from ooodev.adapter.container.index_container_comp import IndexContainerComp
from ooodev.office import draw as mDraw
from ooodev.proto.component_proto import ComponentT
from ooodev.utils import color as mColor
from ooodev.utils import lo as mLo
from ooodev.utils.data_type.image_offset import ImageOffset
from ooodev.utils.data_type.poly_sides import PolySides
from ooodev.utils.dispatch.shape_dispatch_kind import ShapeDispatchKind
from ooodev.utils.kind.drawing_name_space_kind import DrawingNameSpaceKind
from ooodev.utils.kind.drawing_shape_kind import DrawingShapeKind
from ooodev.utils.kind.glue_points_kind import GluePointsKind
from ooodev.utils.kind.presentation_kind import PresentationKind
from ooodev.utils.type_var import PathOrStr
from ooodev.write.write_text import WriteText
from ooodev.exceptions import ex as mEx
from .. import draw_text as mDrawText
from ..shapes import (
    OpenBezierShape,
    ClosedBezierShape,
    DrawShape,
    ConnectorShape,
    EllipseShape,
    OLE2Shape,
    GraphicObjectShape,
    LineShape,
    PolyLineShape,
    PolyPolygonShape,
    RectangleShape,
    TextShape,
)

if TYPE_CHECKING:
    from com.sun.star.animations import XAnimationNode
    from com.sun.star.drawing import GluePoint2
    from com.sun.star.drawing import XShape
    from com.sun.star.text import XText
    from ooo.dyn.presentation.animation_speed import AnimationSpeed
    from ooo.dyn.presentation.fade_effect import FadeEffect
    from ooodev.proto.dispatch_shape import DispatchShape
    from ooodev.units import UnitT
    from ooodev.utils.data_type.size import Size
    from ooodev.utils.kind.drawing_slide_show_kind import DrawingSlideShowKind

_T = TypeVar("_T", bound="ComponentT")


class DrawPagePartial(Generic[_T]):
    """Represents a draw page."""

    # Draw page does implement XDrawPage, but it show in the API of DrawPage Service.

    def __init__(self, owner: _T, component: XDrawPage) -> None:
        self.__owner = owner
        self.__component = component

    def add_bullet(self, bulls_txt: XText, level: int, text: str) -> None:
        """
        Add bullet text to the end of the bullets text area, specifying
        the nesting of the bullet using a numbering level value
        (numbering starts at 0).

        Args:
            bulls_txt (XText): Text object
            level (int): Bullet Level
            text (str): Bullet Text

        Raises:
            DrawError: If error adding bullet.

        Returns:
            None:
        """
        mDraw.Draw.add_bullet(bulls_txt, level, text)

    def add_connector(
        self,
        shape1: XShape,
        shape2: XShape,
        start_conn: GluePointsKind | None = None,
        end_conn: GluePointsKind | None = None,
    ) -> ConnectorShape[_T]:
        """
        Add connector

        Args:
            shape1 (XShape): First Shape to add connector to.
            shape2 (XShape): Second Shape to add connector to.
            start_conn (GluePointsKind | None, optional): Start connector kind. Defaults to right.
            end_conn (GluePointsKind | None, optional): End connector kind. Defaults to left.

        Raises:
            ShapeError: If error occurs.

        Returns:
            XShape: Connector Shape.

        Note:
            Properties for shape can be added or changed by using :py:meth:`~.draw.Draw.set_shape_props`.

            For instance the default value is ``EndShape=ConnectorType.STANDARD``.
            This could be changed.

            .. code-block:: python

                Draw.set_shape_props(shape, EndShape=ConnectorType.CURVE)
        """
        result = mDraw.Draw.add_connector(
            slide=self.component, shape1=shape1, shape2=shape2, start_conn=start_conn, end_conn=end_conn  # type: ignore
        )
        return ConnectorShape(self.__owner, result)

    def add_dispatch_shape(
        self,
        shape_dispatch: ShapeDispatchKind | str,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT,
        fn: DispatchShape,
    ) -> DrawShape[_T]:
        """
        Adds a shape to a Draw slide via a dispatch command

        Args:
            shape_dispatch (ShapeDispatchKind | str): Dispatch Command
            x (int, UnitT): Shape X position in mm units or UnitT.
            y (int, UnitT): Shape Y position in mm units or UnitT
            width (int, UnitT): Shape width in mm units or UnitT
            height (int, UnitT): Shape height in mm units or UnitT
            fn (DispatchShape): Function that is responsible for running the dispatch command and returning the shape.

        Raises:
            NoneError: If adding a dispatch fails.
            ShapeError: If any other error occurs.

        Returns:
            XShape: Shape

        See Also:
            :py:protocol:`~.proto.dispatch_shape.DispatchShape`
        """
        result = mDraw.Draw.add_dispatch_shape(
            slide=self.component, shape_dispatch=shape_dispatch, x=x, y=y, width=width, height=height, fn=fn  # type: ignore
        )
        return DrawShape(self.__owner, result)

    def add_pres_shape(
        self,
        shape_type: PresentationKind,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT,
    ) -> DrawShape[_T]:
        """
        Creates a shape from the "com.sun.star.presentation" package:

        Args:
            shape_type (PresentationKind): Kind of presentation package to create.
            x (int, UnitT): Shape X position in mm units or UnitT.
            y (int, UnitT): Shape Y position in mm units or UnitT.
            width (int, UnitT): Shape width in mm units or UnitT.
            height (int, UnitT): Shape height in mm units or UnitT.

        Raises:
            ShapeError: If error occurs.

        Returns:
            XShape: Presentation Shape.
        """
        result = mDraw.Draw.add_pres_shape(
            slide=self.component, shape_type=shape_type, x=x, y=y, width=width, height=height  # type: ignore
        )
        return DrawShape(self.__owner, result)

    def add_shape(
        self,
        shape_type: DrawingShapeKind | str,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT,
    ) -> DrawShape[_T]:
        """
        Adds a shape to a slide.

        Args:
            shape_type (DrawingShapeKind | str): Shape type.
            x (int, UnitT): Shape X position in mm units or UnitT.
            y (int, UnitT): Shape Y position in mm units or UnitT.
            width (int, UnitT): Shape width in mm units or UnitT.
            height (int, UnitT): Shape height in mm units or UnitT.

        Raises:
            ShapeError: If error occurs.

        Returns:
            XShape: Newly added Shape.

        See Also:
            - :py:meth:`~.draw.Draw.warns_position`
            - :py:meth:`~.draw.Draw.make_shape`
        """
        result = mDraw.Draw.add_shape(slide=self.component, shape_type=shape_type, x=x, y=y, width=width, height=height)  # type: ignore
        return DrawShape(self.__owner, result)

    def add_slide_number(self) -> DrawShape[_T]:
        """
        Adds slide number to a slide

        Args:
            slide (XDrawPage): Slide

        Raises:
            ShapeError: If error occurs.

        Returns:
            DrawShape: Slide number shape.
        """
        result = mDraw.Draw.add_slide_number(slide=self.component)  # type: ignore
        return DrawShape(self.__owner, result)

    def blank_slide(self) -> None:
        """
        Inserts a blank slide

        Args:
            slide (XDrawPage): Slide

        Raises:
            DrawError: If error occurs

        Returns:
            None:
        """
        mDraw.Draw.blank_slide(self.component)  # type: ignore

    def bullets_slide(self, title: str) -> mDrawText.DrawText[_T]:
        """
        Add text to the slide page by treating it as a bullet page, which
        has two text shapes: one for the title, the other for a sequence of
        bullet points; add the title text but return a reference to the bullet
        text area

        Args:
            slide (XDrawPage): Slide
            title (str): Title

        Raises:
            DrawError: If error setting slide.

        Returns:
            DrawText: Text Object
        """
        result = mDraw.Draw.bullets_slide(self.component, title)  # type: ignore
        return mDrawText.DrawText(self.__owner, result)

    def copy_shape(self, old_shape: XShape) -> DrawShape[_T]:
        """
        Copies a shape

        Args:
            old_shape (XShape): Old Shape

        Raises:
            ShapeError: If unable to copy shape.

        Returns:
            DrawShape: Newly Copied shape.
        """
        result = mDraw.Draw.copy_shape(self.component, old_shape)  # type: ignore
        return DrawShape(self.__owner, result)

    def copy_shape_contents(self, old_shape: XShape) -> DrawShape[_T]:
        """
        Copies a shapes contents from old shape into new shape.

        Args:
            old_shape (XShape): Old shape

        Raises:
            ShapeError: If unable to copy shape contents.

        Returns:
            DrawShape: New shape with contents of old shape copied.
        """
        result = mDraw.Draw.copy_shape_contents(self.component, old_shape)  # type: ignore
        return DrawShape(self.__owner, result)

    def draw_bezier_open(self, pts: Sequence[Point], flags: Sequence[PolygonFlags]) -> OpenBezierShape[_T]:
        """
        Draws a bezier curve.

        Args:
            pts (Sequence[Point]): Points
            flags (Sequence[PolygonFlags]): Flags

        Raises:
            IndexError: If ``pts`` and ``flags`` do not have the same number of elements.
            ShapeError: If unable to create Bezier Shape.

        Returns:
            OpenBezierShape: Bezier Shape.
        """
        shape = mDraw.Draw.draw_bezier(slide=self.component, pts=pts, flags=flags, is_open=True)  # type: ignore
        return OpenBezierShape(self.__owner, shape)

    def draw_bezier_closed(self, pts: Sequence[Point], flags: Sequence[PolygonFlags]) -> ClosedBezierShape[_T]:
        """
        Draws a bezier curve.

        Args:
            pts (Sequence[Point]): Points
            flags (Sequence[PolygonFlags]): Flags

        Raises:
            IndexError: If ``pts`` and ``flags`` do not have the same number of elements.
            ShapeError: If unable to create Bezier Shape.

        Returns:
            ClosedBezierShape: Bezier Shape.
        """
        shape = mDraw.Draw.draw_bezier(slide=self.component, pts=pts, flags=flags, is_open=False)  # type: ignore
        return ClosedBezierShape(self.__owner, shape)

    def draw_circle(self, x: int | UnitT, y: int | UnitT, radius: int | UnitT) -> EllipseShape[_T]:
        """
        Gets a circle

        Args:
            x (int, UnitT): Shape X position in mm units or UnitT.
            y (int, UnitT): Shape Y position in mm units or UnitT.
            radius (int, UnitT): Shape radius in mm units or UnitT.

        Raises:
            ShapeError: If error occurs.

        Returns:
            EllipseShape: Circle Shape.
        """
        shape = mDraw.Draw.draw_circle(slide=self.component, x=x, y=y, radius=radius)  # type: ignore
        return EllipseShape(self.__owner, shape)

    def draw_ellipse(
        self, x: int | UnitT, y: int | UnitT, width: int | UnitT, height: int | UnitT
    ) -> EllipseShape[_T]:
        """
        Gets an ellipse

        Args:
            x (int, UnitT): Shape X position in mm units or UnitT.
            y (int, UnitT): Shape Y position in mm units or UnitT.
            width (int, UnitT): Shape width in mm units or UnitT.
            height (int, UnitT): Shape height in mm units or UnitT.

        Raises:
            ShapeError: If error occurs.

        Returns:
            EllipseShape: Ellipse Shape.
        """
        shape = mDraw.Draw.draw_ellipse(slide=self.component, x=x, y=y, width=width, height=height)  # type: ignore
        return EllipseShape(self.__owner, shape)

    def draw_formula(
        self, formula: str, x: int | UnitT, y: int | UnitT, width: int | UnitT, height: int | UnitT
    ) -> OLE2Shape:
        """
        Draws a formula

        Args:
            slide (XDrawPage): Slide
            formula (str): Formula as string to draw/
            x (int, UnitT): Shape X position in mm units or UnitT.
            y (int, , UnitT): Shape Y position in mm units or UnitT
            width (int, , UnitT): Shape width in mm units or UnitT
            height (int, , UnitT): Shape height in mm units or UnitT

        Raises:
            ShapeError: If error occurs.

        Returns:
            XShape: Formula Shape.
        """
        shape = mDraw.Draw.draw_formula(
            slide=self.component, formula=formula, x=x, y=y, width=width, height=height  # type: ignore
        )
        return OLE2Shape(self, shape)

    # region draw_image()
    @overload
    def draw_image(self, fnm: PathOrStr) -> GraphicObjectShape[_T]:
        """
        Draws an image.

        Args:
            fnm (PathOrStr): Path to image

        Returns:
            GraphicObjectShape: Shape
        """
        ...

    @overload
    def draw_image(self, fnm: PathOrStr, x: int | UnitT, y: int | UnitT) -> GraphicObjectShape[_T]:
        """
        Draws an image.

        Args:
            fnm (PathOrStr): Path to image
            x (int, UnitT): Shape X position in mm units or UnitT.
            y (int, UnitT): Shape Y position in mm units or UnitT.

        Returns:
            GraphicObjectShape: Shape
        """
        ...

    @overload
    def draw_image(
        self, fnm: PathOrStr, x: int | UnitT, y: int | UnitT, width: int | UnitT, height: int | UnitT
    ) -> GraphicObjectShape[_T]:
        """
        Draws an image.

        Args:
            fnm (PathOrStr): Path to image
            x (int, UnitT): Shape X position in mm units or UnitT.
            y (int, UnitT): Shape Y position in mm units or UnitT.
            width (int, UnitT): Shape width in mm units or UnitT.
            height (int, UnitT): Shape height in mm units or UnitT.

        Returns:
            GraphicObjectShape: Shape
        """
        ...

    def draw_image(self, *args, **kwargs) -> GraphicObjectShape[_T]:
        """
        Draws an image.

        Args:
            fnm (PathOrStr): Path to image
            x (int, UnitT): Shape X position in mm units or UnitT.
            y (int, UnitT): Shape Y position in mm units or UnitT.
            width (int, UnitT): Shape width in mm units or UnitT.
            height (int, UnitT): Shape height in mm units or UnitT.

        Raises:
            ShapeError: If error occurs.

        Returns:
            GraphicObjectShape: Shape
        """
        shape = mDraw.Draw.draw_image(self.component, *args, **kwargs)  # type: ignore
        return GraphicObjectShape(self.__owner, shape)

    # endregion draw_image()

    def draw_image_offset(
        self, fnm: PathOrStr, xoffset: ImageOffset | float, yoffset: ImageOffset | float
    ) -> GraphicObjectShape[_T] | None:
        """
        Insert the specified picture onto the slide page in the doc
        presentation document. Use the supplied (x, y) offsets to locate the
        top-left of the image.

        Args:
            slide (XDrawPage): Slide
            fnm (PathOrStr): Path to image.
            xoffset (ImageOffset, float): X Offset with value between ``0.0`` and ``1.0``
            yoffset (ImageOffset, float): Y Offset with value between ``0.0`` and ``1.0``

        Returns:
            GraphicObjectShape | None: Shape on success, ``None`` otherwise.
        """
        shape = mDraw.Draw.draw_image_offset(slide=self.component, fnm=fnm, xoffset=xoffset, yoffset=yoffset)  # type: ignore
        if shape is None:
            return None
        return GraphicObjectShape(self.__owner, shape)

    def draw_line(self, x1: int | UnitT, y1: int | UnitT, x2: int | UnitT, y2: int | UnitT) -> LineShape[_T]:
        """
        Draws a line.

        Args:
            x1 (int, UnitT): Line start X position in mm units or UnitT.
            y1 (int, UnitT): Line start Y position mm units or UnitT.
            x2 (int, UnitT): Line end X position mm units or UnitT.
            y2 (int, UnitT): Line end Y position mm units or UnitT.

        Raises:
            ValueError: If x values and y values are a point and not a line.
            ShapeError: If unable to create Line.

        Returns:
            LineShape: Line Shape.
        """
        shape = mDraw.Draw.draw_line(slide=self.component, x1=x1, y1=y1, x2=x2, y2=y2)  # type: ignore
        return LineShape(self.__owner, shape)

    def draw_lines(self, xs: Sequence[Union[int, UnitT]], ys: Sequence[Union[int, UnitT]]) -> PolyLineShape[_T]:
        """
        Draw lines

        Args:
            xs (Sequence[Union[int, UnitT]): Sequence of X positions in mm units or UnitT.
            ys (Sequence[Union[int, UnitT]): Sequence of Y positions in mm units or UnitT.

        Raises:
            IndexError: If ``xs`` and ``xy`` do not have the same number of elements.
            ShapeError: If any other error occurs.

        Returns:
            PolyLineShape: Lines Shape.

        Note:
            The number of points must be the same for both ``xs`` and ``ys``.
        """
        shape = mDraw.Draw.draw_lines(slide=self.component, xs=xs, ys=ys)  # type: ignore
        return PolyLineShape(self.__owner, shape)

    def draw_media(
        self, fnm: PathOrStr, x: int | UnitT, y: int | UnitT, width: int | UnitT, height: int | UnitT
    ) -> DrawShape[_T]:
        """
        Draws media.

        Args:
            fnm (PathOrStr): Path to Media file.
            x (int, UnitT): Shape X position in mm units or UnitT.
            y (int, UnitT): Shape Y position in mm units or UnitT.
            width (int, UnitT): Shape width in mm units or UnitT.
            height (int, UnitT): Shape height in mm units or UnitT.

        Raises:
            ShapeError: If error occurs.

        Returns:
            XShape: Media shape.
        """
        # could not find MediaShape in api.
        # https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1drawing.html
        # however it can be found in examples.
        # https://ask.libreoffice.org/t/how-to-add-video-to-impress-with-python/33050/2
        shape = mDraw.Draw.draw_media(slide=self.component, fnm=fnm, x=x, y=y, width=width, height=height)  # type: ignore
        return DrawShape(self.__owner, shape)

    def draw_polar_line(self, x: int | UnitT, y: int | UnitT, degrees: int, distance: int | UnitT) -> LineShape[_T]:
        """
        Draw a line from ``x``, ``y`` in the direction of degrees, for the specified distance
        degrees is measured clockwise from x-axis

        Args:
            x (int, UnitT): Shape X position in mm units or UnitT.
            y (int, UnitT): Shape Y position in mm units or UnitT.
            degrees (int): Direction of degrees
            distance (int, UnitT): Distance of line in mm units or UnitT..

        Raises:
            ShapeError: If unable to create Polar Line Shape.

        Returns:
            LineShape: Polar Line Shape.
        """
        shape = mDraw.Draw.draw_polar_line(slide=self.component, x=x, y=y, degrees=degrees, distance=distance)  # type: ignore
        return LineShape(self.__owner, shape)

    # region draw_polygon()
    @overload
    def draw_polygon(self, x: int | UnitT, y: int | UnitT, sides: PolySides | int) -> PolyPolygonShape[_T]:
        """
        Gets a polygon.

        Args:
            x (int, UnitT): Shape X position in mm units or UnitT.
            y (int, UnitT): Shape Y position in mm units or UnitT.
            sides (PolySides | int): Polygon Sides value from ``3`` to ``30``.

        Returns:
            PolyPolygonShape: Polygon Shape.
        """
        ...

    @overload
    def draw_polygon(
        self, x: int | UnitT, y: int | UnitT, sides: PolySides | int, radius: int
    ) -> PolyPolygonShape[_T]:
        """
        Gets a polygon.

        Args:
            x (int, UnitT): Shape X position in mm units or UnitT.
            y (int, UnitT): Shape Y position in mm units or UnitT.
            sides (PolySides | int): Polygon Sides value from ``3`` to ``30``.
            radius (int, optional): Shape radius in mm units. Defaults to the value of :py:attr:`draw.Draw.POLY_RADIUS`.

        Returns:
            PolyPolygonShape: Polygon Shape.
        """
        ...

    def draw_polygon(
        self, x: int | UnitT, y: int | UnitT, sides: PolySides | int, radius: int = mDraw.Draw.POLY_RADIUS
    ) -> PolyPolygonShape[_T]:
        """
        Gets a polygon.

        Args:
            x (int, UnitT): Shape X position in mm units or UnitT.
            y (int, UnitT): Shape Y position in mm units or UnitT.
            sides (PolySides | int): Polygon Sides value from ``3`` to ``30``.
            radius (int, optional): Shape radius in mm units. Defaults to the value of :py:attr:`draw.Draw.POLY_RADIUS`.

        Raises:
            ShapeError: If error occurs.

        Returns:
            PolyPolygonShape: Polygon Shape.
        """
        shape = mDraw.Draw.draw_polygon(slide=self.component, x=x, y=y, sides=sides, radius=radius)  # type: ignore
        return PolyPolygonShape(self.__owner, shape)

    # endregion draw_polygon()

    def draw_rectangle(
        self, x: int | UnitT, y: int | UnitT, width: int | UnitT, height: int | UnitT
    ) -> RectangleShape[_T]:
        """
        Gets a rectangle.

        Args:
            x (int, UnitT): Shape X position in mm units or UnitT.
            y (int, UnitT): Shape Y position in mm units or UnitT.
            width (int, UnitT): Shape width in mm units or UnitT.
            height (int, UnitT): Shape height in mm units or UnitT.

        Raises:
            ShapeError: If error occurs.

        Returns:
            RectangleShape: Rectangle Shape.
        """
        shape = mDraw.Draw.draw_rectangle(slide=self.component, x=x, y=y, width=width, height=height)  # type: ignore
        return RectangleShape(self.__owner, shape)

    def draw_text(
        self,
        msg: str,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT,
        font_size: float | UnitT = 0,
    ) -> TextShape[_T]:
        """
        Draws Text.

        Args:
            slide (XDrawPage): Slide.
            msg (str): Text to draw.
            x (int, UnitT): Shape X position in mm units or UnitT.
            y (int, UnitT): Shape Y position in mm units or UnitT.
            width (int, UnitT): Shape width in mm units or UnitT.
            height (int, UnitT): Shape height in mm units or UnitT.
            font_size (float, UnitT, optional): Font size of text in Points or UnitT.

        Raises:
            ShapeError: If error occurs.

        Returns:
            TextShape: Shape
        """
        shape = mDraw.Draw.draw_text(
            slide=self.component, msg=msg, x=x, y=y, width=width, height=height, font_size=font_size  # type: ignore
        )
        return TextShape(self.__owner, shape)

    def find_biggest_zorder(self) -> int:
        """
        Finds the shape with the largest z-order.

        Raises:
            DrawError: If unable to find biggest z-order.

        Returns:
            int: Z-Order
        """
        return mDraw.Draw.find_biggest_zorder(self.component)  # type: ignore

    def find_shape_by_name(self, shape_name: str) -> DrawShape[_T]:
        """
        Finds a shape by its name.

        Args:
            shape_name (str): Shape Name.

        Raise:
            ShapeMissingError: If shape is not found.
            ShapeError: If any other error occurs.

        Returns:
            DrawShape: Shape.
        """
        shape = mDraw.Draw.find_shape_by_name(self.component, shape_name)  # type: ignore
        return DrawShape(self.__owner, shape)

    def find_shape_by_type(self, shape_type: DrawingNameSpaceKind | str) -> DrawShape[_T]:
        """
        Finds a shape by its type

        Args:
            shape_type (DrawingNameSpaceKind | str): Shape Type

        Raise:
            ShapeMissingError: If shape is not found.
            ShapeError: If any other error occurs.

        Returns:
            DrawShape: Shape
        """
        shape = mDraw.Draw.find_shape_by_type(self.component, shape_type)  # type: ignore
        return DrawShape(self.__owner, shape)

    def find_top_shape(self) -> DrawShape[_T]:
        """
        Gets the top most shape of a slide.

        Raises:
            ShapeMissingError: If there are no shapes for slide or unable to find top shape.
            ShapeError: If any other error occurs.

        Returns:
            DrawShape: Top most shape.
        """
        shape = mDraw.Draw.find_top_shape(self.component)  # type: ignore
        return DrawShape(self.__owner, shape)

    def get_animation_node(self) -> XAnimationNode:
        """
        Gets Animation Node

        Args:
            slide (XDrawPage): Slide

        Raises:
            DrawPageError: If error occurs.

        Returns:
            XAnimationNode: Animation Node
        """
        return mDraw.Draw.get_animation_node(self.component)  # type: ignore

    def title_only_slide(self, header: str) -> None:
        """
        Creates a slide with only a title

        Args:
            header (str): Header text.

        Raises:
            DrawError: If error occurs.

        Returns:
            None:
        """
        mDraw.Draw.title_only_slide(self.component, header)  # type: ignore

    def get_chart_shape(
        self, x: int | UnitT, y: int | UnitT, width: int | UnitT, height: int | UnitT
    ) -> OLE2Shape[_T]:
        """
        Gets a chart shape.

        Args:
            x (int, UnitT): Shape X position in mm units or UnitT.
            y (int, UnitT): Shape Y position in mm units or UnitT.
            width (int, UnitT): Shape width in mm units or UnitT.
            height (int, UnitT): Shape height in mm units or UnitT.

        Raises:
            ShapeError: If Error occurs.

        Returns:
            OLE2Shape: Chart Shape.
        """
        shape = mDraw.Draw.get_chart_shape(slide=self.component, x=x, y=y, width=width, height=height)  # type: ignore
        return OLE2Shape(self.__owner, shape)

    def get_fill_color(self) -> mColor.Color:
        """
        Gets the fill color of a shape.

        Args:
            shape (XShape): Shape

        Raises:
            ColorError: If error occurs.

        Returns:
            ~ooodev.utils.color.Color: Color
        """
        return mDraw.Draw.get_fill_color(self.component)  # type: ignore

    def get_form_container(self) -> IndexContainerComp | None:
        """
        Gets form container.
        The first form in slide is returned if found.

        Raises:
            DrawError: If error occurs.

        Returns:
            IndexContainerComp | None: Form Container on success, None otherwise.
        """
        container = mDraw.Draw.get_form_container(self.component)  # type: ignore
        if container is None:
            return None
        return IndexContainerComp(container)

    def get_glue_points(self) -> Tuple[GluePoint2, ...]:
        """
        Gets Glue Points

        Args:
            shape (XShape): Shape

        Raises:
            DrawError: If error occurs

        Returns:
            Tuple[GluePoint2, ...]: Glue Points.

        Note:
            If a glue point can not be accessed then it is ignored.
        """
        return mDraw.Draw.get_glue_points(self.component)  # type: ignore

    def get_line_color(self) -> mColor.Color:
        """
        Gets the line color of a shape.

        Raises:
            ColorError: If error occurs.

        Returns:
            ~ooodev.utils.color.Color: Color
        """
        return mDraw.Draw.get_line_color(self.component)  # type: ignore

    def get_line_thickness(self) -> int:
        """
        Gets line thickness of a shape.

        Returns:
            int: Line Thickness on success; Otherwise, ``0``.
        """
        return mDraw.Draw.get_line_thickness(self.component)  # type: ignore

    def get_ordered_shapes(self) -> List[DrawShape[_T]]:
        """
        Gets ordered shapes

        Returns:
            List[DrawShape[_T]]: List of Ordered Shapes.

        See Also:
            :py:meth:`~.draw.Draw.get_shapes`
        """
        shapes = mDraw.Draw.get_ordered_shapes(slide=self.component)  # type: ignore
        return [DrawShape(self.__owner, shape) for shape in shapes]

    def get_shape_text(self) -> str:
        """
        Gets the text from inside a shape.

        Raises:
            DrawError: If error occurs getting shape text.

        Returns:
            str: Shape text
        """
        return mDraw.Draw.get_shape_text(slide=self.__component)

    def get_shapes(self) -> List[DrawShape[_T]]:
        """
        Gets shapes

        Returns:
            List[DrawShape[_T]]: List of Shapes.

        See Also:
            :py:meth:`~.draw.Draw.get_ordered_shapes`
        """
        shapes = mDraw.Draw.get_shapes(slide=self.component)  # type: ignore
        return [DrawShape(self.__owner, shape) for shape in shapes]

    def get_size_mm(self) -> Size:
        """
        Gets size of the given slide page (in mm units)

        Raises:
            SizeError: If unable to get size.

        Returns:
            ~ooodev.utils.data_type.size.Size: Size struct.
        """
        return mDraw.Draw.get_slide_size(self.__component)

    def get_slide_number(self) -> int:
        """
        Gets slide number.

        Raises:
            DrawError: If error occurs.

        Returns:
            int: Slide Number.
        """
        return mDraw.Draw.get_slide_number(slide=self.component)  # type: ignore

    def get_slide_title(self) -> str | None:
        """
        Gets slide title if it exist.

        Raises:
            DrawError: If error occurs.

        Returns:
            str | None: Slide Title on success; Otherwise, ``None``.
        """
        return mDraw.Draw.get_slide_title(slide=self.component)  # type: ignore

    def goto_page(self) -> None:
        """
        Go to page represented by this object.

        Raises:
            DrawError: If error occurs.

        Returns:
            None:
        """
        if not self.__owner:
            raise mEx.DrawPageError("DrawPage owner is None")
        if not mLo.Lo.is_uno_interfaces(self.__owner, "com.sun.star.drawing.XDrawPage"):
            raise mEx.DrawPageError("DrawPage component is not XDrawPage")
        mDraw.Draw.goto_page(doc=self.__owner, page=self.component)  # type: ignore

    def move_to_bottom(self, shape: XShape) -> None:
        """
        Moves a shape to the bottom of the z-order

        Args:
            shape (XShape): Shape

        Raises:
            ShapeMissingError: If unable to find shapes for slide.
            ShapeError: If any other error occurs.

        Returns:
            None:
        """
        mDraw.Draw.move_to_bottom(slide=self.component, shape=shape)  # type: ignore

    def move_to_top(self, shape: XShape) -> None:
        """
        Moves a shape to the top of the z-order

        Args:
            shape (XShape): Shape

        Raises:
            ShapeMissingError: If unable to find shapes for slide.
            ShapeError: If any other error occurs.

        Returns:
            None:
        """
        mDraw.Draw.move_to_top(slide=self.component, shape=shape)  # type: ignore

    def remove_master_page(self) -> None:
        """
        Removes a master page

        Args:
            slide (XDrawPage): Draw page to remove

        Raises:
            DrawError: If unable to remove master page/

        Returns:
            None:
        """
        if not self.__owner:
            raise mEx.DrawPageError("Owner is None")
        if not mLo.Lo.is_uno_interfaces(self.__owner, XDrawPage):
            raise mEx.DrawPageError("Owner component is not XDrawPage")
        mDraw.Draw.remove_master_page(doc=self.__owner, slide=self.__component)  # type: ignore

    def save_page(self, fnm: PathOrStr, mime_type: str) -> None:
        """
        Saves a Draw page to file.

        Args:
            fnm (PathOrStr): Path to save page as
            mime_type (str): Mime Type of page to save as such as ``image/jpeg`` or ``image/png``.

        Raises:
            DrawError: If error occurs.

        Returns:
            None:

        See Also:
            :py:meth:`ooodev.utils.images_lo.ImagesLo.change_to_mime`.
        """
        mDraw.Draw.save_page(self.__component, fnm, mime_type)

    def set_name(self, name: str) -> None:
        """
        Sets the name of a slide.

        Args:
            name (str): Name.

        Raises:
            DrawError: If error occurs setting name.

        Returns:
            None:
        """
        mDraw.Draw.set_name(slide=self.component, name=name)  # type: ignore

    def set_transition(
        self,
        fade_effect: FadeEffect,
        speed: AnimationSpeed,
        change: DrawingSlideShowKind,
        duration: int,
    ) -> None:
        """
        Sets the transition for a slide.

        Args:
            slide (XDrawPage): Slide
            fade_effect (FadeEffect): Fade Effect
            speed (AnimationSpeed): Animation Speed
            change (SlideShowKind): Slide show kind
            duration (int): Duration of slide. Only used when ``change=SlideShowKind.AUTO_CHANGE``

        Raises:
            DrawPageError: If error occurs.

        Returns:
            None:
        """
        mDraw.Draw.set_transition(
            slide=self.component, fade_effect=fade_effect, speed=speed, change=change, duration=duration  # type: ignore
        )
