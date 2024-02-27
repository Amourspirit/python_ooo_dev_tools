from __future__ import annotations
from typing import List, Tuple, overload, Sequence, TYPE_CHECKING, TypeVar, Generic, Union
import uno


from com.sun.star.drawing import XDrawPage

from ooo.dyn.awt.point import Point
from ooo.dyn.drawing.polygon_flags import PolygonFlags
from ooodev.mock import mock_g
from ooodev.adapter.container.index_container_comp import IndexContainerComp
from ooodev.exceptions import ex as mEx
from ooodev.office import draw as mDraw
from ooodev.utils import color as mColor
from ooodev.loader import lo as mLo
from ooodev.utils.context.lo_context import LoContext
from ooodev.utils.data_type.image_offset import ImageOffset
from ooodev.utils.data_type.poly_sides import PolySides
from ooodev.utils.dispatch.shape_dispatch_kind import ShapeDispatchKind
from ooodev.utils.kind.drawing_name_space_kind import DrawingNameSpaceKind
from ooodev.utils.kind.drawing_shape_kind import DrawingShapeKind
from ooodev.utils.kind.glue_points_kind import GluePointsKind
from ooodev.utils.kind.presentation_kind import PresentationKind
from ooodev.proto.component_proto import ComponentT

# more import at bottom of this module

if TYPE_CHECKING:
    from com.sun.star.animations import XAnimationNode
    from com.sun.star.drawing import GluePoint2
    from com.sun.star.drawing import XShape
    from com.sun.star.text import XText
    from ooo.dyn.presentation.animation_speed import AnimationSpeed
    from ooo.dyn.presentation.fade_effect import FadeEffect
    from ooodev.draw.shapes import ShapeBase
    from ooodev.proto.dispatch_shape import DispatchShape
    from ooodev.units.unit_obj import UnitT
    from ooodev.utils.data_type.size import Size
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.utils.kind.drawing_slide_show_kind import DrawingSlideShowKind
    from ooodev.utils.type_var import PathOrStr

_T = TypeVar("_T", bound="ComponentT")


class DrawPagePartial(Generic[_T]):
    """Represents a draw page."""

    # Draw page does implement XDrawPage, but it show in the API of DrawPage Service.

    def __init__(self, owner: _T, component: XDrawPage, lo_inst: LoInst | None = None) -> None:
        self.__owner = owner
        self.__component = component
        self.__lo_inst = mLo.Lo.current_lo if lo_inst is None else lo_inst

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
        with LoContext(self.__lo_inst):
            result = mDraw.Draw.add_connector(
                slide=self.__component, shape1=shape1, shape2=shape2, start_conn=start_conn, end_conn=end_conn  # type: ignore
            )
        return ConnectorShape(owner=self.__owner, component=result, lo_inst=self.__lo_inst)

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
        with LoContext(self.__lo_inst):
            result = mDraw.Draw.add_dispatch_shape(
                slide=self.__component, shape_dispatch=shape_dispatch, x=x, y=y, width=width, height=height, fn=fn  # type: ignore
            )
        return DrawShape(owner=self.__owner, component=result, lo_inst=self.__lo_inst)

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
        with LoContext(self.__lo_inst):
            result = mDraw.Draw.add_pres_shape(
                slide=self.__component, shape_type=shape_type, x=x, y=y, width=width, height=height  # type: ignore
            )
        return DrawShape(owner=self.__owner, component=result, lo_inst=self.__lo_inst)

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

        Note:
            If ``x`` or ``y`` is negative or ``0`` then the shape position will not be set.
            If ``width`` or ``height`` is negative or ``0`` then the shape size will not be set.

        .. versionchanged:: 0.17.14
            Now does not set size and/or position unless the values are greater than ``0``.
        """
        with LoContext(self.__lo_inst):
            result = mDraw.Draw.add_shape(slide=self.__component, shape_type=shape_type, x=x, y=y, width=width, height=height)  # type: ignore
        return DrawShape(owner=self.__owner, component=result, lo_inst=self.__lo_inst)

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
        with LoContext(self.__lo_inst):
            result = mDraw.Draw.add_slide_number(slide=self.__component)  # type: ignore
        return DrawShape(owner=self.__owner, component=result, lo_inst=self.__lo_inst)

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
        mDraw.Draw.blank_slide(self.__component)  # type: ignore

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
        result = mDraw.Draw.bullets_slide(self.__component, title)  # type: ignore
        return mDrawText.DrawText(owner=self.__owner, component=result, lo_inst=self.__lo_inst)

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
        with LoContext(self.__lo_inst):
            result = mDraw.Draw.copy_shape(self.__component, old_shape)  # type: ignore
        return DrawShape(owner=self.__owner, component=result, lo_inst=self.__lo_inst)

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
        with LoContext(self.__lo_inst):
            result = mDraw.Draw.copy_shape_contents(self.__component, old_shape)  # type: ignore
        return DrawShape(owner=self.__owner, component=result, lo_inst=self.__lo_inst)

    # region Draw Bezier
    @overload
    def draw_bezier_open(self, pts: Sequence[Point], flags: Sequence[PolygonFlags]) -> OpenBezierShape[_T]:
        """
        Draws a bezier curve.

        Args:
            pts (Sequence[Point]): Points
            flags (Sequence[PolygonFlags]): Flags

        Returns:
            OpenBezierShape: Bezier Shape.
        """
        ...

    @overload
    def draw_bezier_open(
        self, pts: Sequence[Sequence[Point]], flags: Sequence[Sequence[PolygonFlags]]
    ) -> OpenBezierShape[_T]:
        """
        Draws a bezier curve.

        Args:
            pts (Sequence[Sequence[Point]]): Points
            flags (Sequence[Sequence[PolygonFlags]]): Flags

        Returns:
            OpenBezierShape: Bezier Shape.
        """
        ...

    def draw_bezier_open(
        self,
        pts: Sequence[Point] | Sequence[Sequence[Point]],
        flags: Sequence[PolygonFlags] | Sequence[Sequence[PolygonFlags]],
    ) -> OpenBezierShape[_T]:
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
        shape = mDraw.Draw.draw_bezier(slide=self.__component, pts=pts, flags=flags, is_open=True)  # type: ignore
        return OpenBezierShape(owner=self.__owner, component=shape, lo_inst=self.__lo_inst)

    @overload
    def draw_bezier_closed(self, pts: Sequence[Point], flags: Sequence[PolygonFlags]) -> ClosedBezierShape[_T]:
        """
        Draws a bezier curve.

        Args:
            pts (Sequence[Point]): Points
            flags (Sequence[PolygonFlags]): Flags

        Returns:
            ClosedBezierShape: Bezier Shape.
        """
        ...

    @overload
    def draw_bezier_closed(
        self, pts: Sequence[Sequence[Point]], flags: Sequence[Sequence[PolygonFlags]]
    ) -> ClosedBezierShape[_T]:
        """
        Draws a bezier curve.

        Args:
            pts (Sequence[Sequence[Point]]): Points
            flags (Sequence[Sequence[PolygonFlags]]): Flags

        Returns:
            ClosedBezierShape: Bezier Shape.
        """
        ...

    def draw_bezier_closed(
        self,
        pts: Sequence[Point] | Sequence[Sequence[Point]],
        flags: Sequence[PolygonFlags] | Sequence[Sequence[PolygonFlags]],
    ) -> ClosedBezierShape[_T]:
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
        shape = mDraw.Draw.draw_bezier(slide=self.__component, pts=pts, flags=flags, is_open=False)  # type: ignore
        return ClosedBezierShape(owner=self.__owner, component=shape, lo_inst=self.__lo_inst)

    # endregion Draw Bezier

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
        with LoContext(self.__lo_inst):
            shape = mDraw.Draw.draw_circle(slide=self.__component, x=x, y=y, radius=radius)  # type: ignore
        return EllipseShape(owner=self.__owner, component=shape, lo_inst=self.__lo_inst)

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
        with LoContext(self.__lo_inst):
            shape = mDraw.Draw.draw_ellipse(slide=self.__component, x=x, y=y, width=width, height=height)  # type: ignore
        return EllipseShape(owner=self.__owner, component=shape, lo_inst=self.__lo_inst)

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

        Note:
            If ``x`` or ``y`` is negative or ``0`` then the shape position will not be set.
            If ``width`` or ``height`` is negative or ``0`` then the shape size will not be set.

        .. versionchanged:: 0.17.14
            Now does not set size and/or position unless the values are greater than ``0``.
        """
        with LoContext(self.__lo_inst):
            shape = mDraw.Draw.draw_formula(
                slide=self.__component, formula=formula, x=x, y=y, width=width, height=height  # type: ignore
            )
        return OLE2Shape(owner=self.__owner, component=shape, lo_inst=self.__lo_inst)

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
        shape = mDraw.Draw.draw_image(self.__component, *args, **kwargs)  # type: ignore
        return GraphicObjectShape(owner=self.__owner, component=shape, lo_inst=self.__lo_inst)

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
        with LoContext(self.__lo_inst):
            shape = mDraw.Draw.draw_image_offset(slide=self.__component, fnm=fnm, xoffset=xoffset, yoffset=yoffset)  # type: ignore
        if shape is None:
            return None
        return GraphicObjectShape(owner=self.__owner, component=shape, lo_inst=self.__lo_inst)

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
        with LoContext(self.__lo_inst):
            shape = mDraw.Draw.draw_line(slide=self.__component, x1=x1, y1=y1, x2=x2, y2=y2)  # type: ignore
        return LineShape(owner=self.__owner, component=shape, lo_inst=self.__lo_inst)

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
        with LoContext(self.__lo_inst):
            shape = mDraw.Draw.draw_lines(slide=self.__component, xs=xs, ys=ys)  # type: ignore
        return PolyLineShape(owner=self.__owner, component=shape, lo_inst=self.__lo_inst)

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

        Note:
            If ``x`` or ``y`` is negative or ``0`` then the shape position will not be set.
            If ``width`` or ``height`` is negative or ``0`` then the shape size will not be set.

        .. versionchanged:: 0.17.14
            Now does not set size and/or position unless the values are greater than ``0``.
        """
        # could not find MediaShape in api.
        # https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1drawing.html
        # however it can be found in examples.
        # https://ask.libreoffice.org/t/how-to-add-video-to-impress-with-python/33050/2
        with LoContext(self.__lo_inst):
            shape = mDraw.Draw.draw_media(slide=self.__component, fnm=fnm, x=x, y=y, width=width, height=height)  # type: ignore
        return DrawShape(owner=self.__owner, component=shape, lo_inst=self.__lo_inst)

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
        with LoContext(self.__lo_inst):
            shape = mDraw.Draw.draw_polar_line(slide=self.__component, x=x, y=y, degrees=degrees, distance=distance)  # type: ignore
        return LineShape(owner=self.__owner, component=shape, lo_inst=self.__lo_inst)

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
        shape = mDraw.Draw.draw_polygon(slide=self.__component, x=x, y=y, sides=sides, radius=radius)  # type: ignore
        return PolyPolygonShape(owner=self.__owner, component=shape, lo_inst=self.__lo_inst)

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

        Note:
            If ``x`` or ``y`` is negative or ``0`` then the shape position will not be set.
            If ``width`` or ``height`` is negative or ``0`` then the shape size will not be set.

        .. versionchanged:: 0.17.14
            Now does not set size and/or position unless the values are greater than ``0``.
        """
        with LoContext(self.__lo_inst):
            shape = mDraw.Draw.draw_rectangle(slide=self.__component, x=x, y=y, width=width, height=height)  # type: ignore
        return RectangleShape(owner=self.__owner, component=shape, lo_inst=self.__lo_inst)

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

        Note:
            If ``x`` or ``y`` is negative or ``0`` then the shape position will not be set.
            If ``width`` or ``height`` is negative or ``0`` then the shape size will not be set.

        .. versionchanged:: 0.17.14
            Now does not set size and/or position unless the values are greater than ``0``.
        """
        with LoContext(self.__lo_inst):
            shape = mDraw.Draw.draw_text(
                slide=self.__component, msg=msg, x=x, y=y, width=width, height=height, font_size=font_size  # type: ignore
            )
        return TextShape(owner=self.__owner, component=shape, lo_inst=self.__lo_inst)

    def find_biggest_zorder(self) -> int:
        """
        Finds the shape with the largest z-order.

        Raises:
            DrawError: If unable to find biggest z-order.

        Returns:
            int: Z-Order
        """
        return mDraw.Draw.find_biggest_zorder(self.__component)  # type: ignore

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
        shape = mDraw.Draw.find_shape_by_name(self.__component, shape_name)  # type: ignore
        return DrawShape(owner=self.__owner, component=shape, lo_inst=self.__lo_inst)

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
        shape = mDraw.Draw.find_shape_by_type(self.__component, shape_type)  # type: ignore
        return DrawShape(owner=self.__owner, component=shape, lo_inst=self.__lo_inst)

    def find_top_shape(self) -> DrawShape[_T]:
        """
        Gets the top most shape of a slide.

        Raises:
            ShapeMissingError: If there are no shapes for slide or unable to find top shape.
            ShapeError: If any other error occurs.

        Returns:
            DrawShape: Top most shape.
        """
        shape = mDraw.Draw.find_top_shape(self.__component)  # type: ignore
        return DrawShape(owner=self.__owner, component=shape, lo_inst=self.__lo_inst)

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
        result = mDraw.Draw.get_animation_node(self.__component)  # type: ignore
        return result

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
        mDraw.Draw.title_only_slide(self.__component, header)  # type: ignore

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
        with LoContext(self.__lo_inst):
            shape = mDraw.Draw.get_chart_shape(slide=self.__component, x=x, y=y, width=width, height=height)  # type: ignore
        return OLE2Shape(owner=self.__owner, component=shape, lo_inst=self.__lo_inst)

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
        return mDraw.Draw.get_fill_color(self.__component)  # type: ignore

    def get_form_container(self) -> IndexContainerComp | None:
        """
        Gets form container.
        The first form in slide is returned if found.

        Raises:
            DrawError: If error occurs.

        Returns:
            IndexContainerComp | None: Form Container on success, None otherwise.
        """
        container = mDraw.Draw.get_form_container(self.__component)  # type: ignore
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
        return mDraw.Draw.get_glue_points(self.__component)  # type: ignore

    def get_line_color(self) -> mColor.Color:
        """
        Gets the line color of a shape.

        Raises:
            ColorError: If error occurs.

        Returns:
            ~ooodev.utils.color.Color: Color
        """
        return mDraw.Draw.get_line_color(self.__component)  # type: ignore

    def get_line_thickness(self) -> int:
        """
        Gets line thickness of a shape.

        Returns:
            int: Line Thickness on success; Otherwise, ``0``.
        """
        return mDraw.Draw.get_line_thickness(self.__component)  # type: ignore

    def get_name(self) -> str:
        """
        Gets the name of the slide.

        Raises:
            DrawError: If error occurs setting name.

        Returns:
            str: Slide name.

        .. versionadded:: 0.17.13
        """
        return mDraw.Draw.get_name(slide=self.__component)  # type: ignore

    def get_ordered_shapes(self) -> List[DrawShape[_T]]:
        """
        Gets ordered shapes

        Returns:
            List[DrawShape[_T]]: List of Ordered Shapes.

        See Also:
            :py:meth:`~.draw.Draw.get_shapes`
        """
        shapes = mDraw.Draw.get_ordered_shapes(slide=self.__component)  # type: ignore
        return [DrawShape(owner=self.__owner, component=shape, lo_inst=self.__lo_inst) for shape in shapes]

    def get_shape_text(self) -> str:
        """
        Gets the text from inside a shape.

        Raises:
            DrawError: If error occurs getting shape text.

        Returns:
            str: Shape text
        """
        return mDraw.Draw.get_shape_text(slide=self.__component)

    def get_shapes(self) -> List[ShapeBase[_T]]:
        """
        Gets shapes

        Returns:
            List[ShapeBase[_T]]: List of Shapes.

        See Also:
            :py:meth:`~.draw.Draw.get_ordered_shapes`
        """
        # pylint: disable=import-outside-toplevel
        shapes = mDraw.Draw.get_shapes(slide=self.__component)  # type: ignore
        from ooodev.draw.shapes.partial.shape_factory_partial import ShapeFactoryPartial

        factory = ShapeFactoryPartial(owner=self.__owner, lo_inst=self.__lo_inst)
        return [factory.shape_factory(shape) for shape in shapes]

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
        return mDraw.Draw.get_slide_number(slide=self.__component)  # type: ignore

    def get_slide_title(self) -> str | None:
        """
        Gets slide title if it exist.

        Raises:
            DrawError: If error occurs.

        Returns:
            str | None: Slide Title on success; Otherwise, ``None``.
        """
        return mDraw.Draw.get_slide_title(slide=self.__component)  # type: ignore

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
        if not self.__lo_inst.is_uno_interfaces(self.__owner, "com.sun.star.drawing.XDrawPage"):
            raise mEx.DrawPageError("DrawPage component is not XDrawPage")
        mDraw.Draw.goto_page(doc=self.__owner, page=self.__component)  # type: ignore

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
        mDraw.Draw.move_to_bottom(slide=self.__component, shape=shape)  # type: ignore

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
        mDraw.Draw.move_to_top(slide=self.__component, shape=shape)  # type: ignore

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
        if not self.__lo_inst.is_uno_interfaces(self.__owner, XDrawPage):
            raise mEx.DrawPageError("Owner component is not XDrawPage")
        mDraw.Draw.remove_master_page(doc=self.__owner, slide=self.__component)  # type: ignore

    def save_page(self, fnm: PathOrStr, mime_type: str, filter_data: dict | None = None) -> None:
        """
        Saves a Draw page to file.

        Args:
            fnm (PathOrStr): Path to save page as
            mime_type (str): Mime Type of page to save as such as ``image/jpeg`` or ``image/png``.
            filter_data (dict, optional): Filter data. Defaults to ``None``.

        Raises:
            DrawError: If error occurs.

        Returns:
            None:

        See Also:
            - :ref:`ooodev.draw.filter.export_png`
            - :ref:`ooodev.draw.filter.export_jpg`
            - :py:meth:`ImagesLo.change_to_mime() <ooodev.utils.images_lo.ImagesLo.change_to_mime>`.
            - :py:meth:`ImagesLo.get_dpi_width_height() <ooodev.utils.images_lo.ImagesLo.get_dpi_width_height>`.

        .. versionchanged:: 0.21.3
            Added `filter_data` parameter.
        """
        with LoContext(self.__lo_inst):
            mDraw.Draw.save_page(self.__component, fnm, mime_type, filter_data)

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
        mDraw.Draw.set_name(slide=self.__component, name=name)  # type: ignore

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
            slide=self.__component, fade_effect=fade_effect, speed=speed, change=change, duration=duration  # type: ignore
        )


# These import have to be here to avoid circular imports.
# pylint: disable=wrong-import-position
# ruff: noqa: E402

from ooodev.draw import draw_text as mDrawText
from ooodev.draw.shapes.open_bezier_shape import OpenBezierShape
from ooodev.draw.shapes.closed_bezier_shape import ClosedBezierShape
from ooodev.draw.shapes.draw_shape import DrawShape
from ooodev.draw.shapes.connector_shape import ConnectorShape
from ooodev.draw.shapes.ellipse_shape import EllipseShape
from ooodev.draw.shapes.ole2_shape import OLE2Shape
from ooodev.draw.shapes.graphic_object_shape import GraphicObjectShape
from ooodev.draw.shapes.line_shape import LineShape
from ooodev.draw.shapes.poly_line_shape import PolyLineShape
from ooodev.draw.shapes.poly_polygon_shape import PolyPolygonShape
from ooodev.draw.shapes.rectangle_shape import RectangleShape
from ooodev.draw.shapes.text_shape import TextShape

if mock_g.FULL_IMPORT:
    from ooodev.draw.shapes.partial.shape_factory_partial import ShapeFactoryPartial
