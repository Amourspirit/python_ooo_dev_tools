from __future__ import annotations
from typing import TYPE_CHECKING, TypeVar, Generic, overload, Tuple
import uno
from com.sun.star.drawing import XDrawPage
from com.sun.star.text import XText
from ooodev.exceptions import ex as mEx
from ooodev.utils import lo as mLo
from ooodev.proto.component_proto import ComponentT
from ooodev.office import draw as mDraw
from ooodev.utils.data_type.angle import Angle
from ooodev.utils.kind.drawing_bitmap_kind import DrawingBitmapKind
from ooodev.utils.kind.drawing_gradient_kind import DrawingGradientKind
from ooodev.utils.kind.drawing_hatching_kind import DrawingHatchingKind
from .shape_text_cursor import ShapeTextCursor


_T = TypeVar("_T", bound="ComponentT")

if TYPE_CHECKING:
    from com.sun.star.awt import Gradient
    from com.sun.star.awt import Point
    from com.sun.star.beans import XPropertySet
    from com.sun.star.container import XNameContainer
    from com.sun.star.drawing import GluePoint2
    from com.sun.star.drawing import HomogenMatrix3
    from com.sun.star.drawing import XShape
    from com.sun.star.graphic import XGraphic
    from ooo.dyn.drawing.line_style import LineStyle
    from ooodev.utils.type_var import PathOrStr
    from ooodev.utils.data_type.size import Size
    from ooodev.utils import color as mColor
    from ooodev.units import UnitT
    from ooodev.utils.kind.graphic_style_kind import GraphicStyleKind
    from ooodev.proto.size_obj import SizeObj
    from ooodev.utils.data_type.intensity import Intensity


class ShapeBase(
    Generic[_T],
):
    def __init__(self, owner: _T, component: XShape) -> None:
        self.__owner = owner
        self.__component = component

    def get_shape_text_cursor(self) -> ShapeTextCursor[_T]:
        """
        Gets a cursor object for this text.

        Returns:
            ShapeTextCursor: Cursor.
        """
        xtext = mLo.Lo.qi(XText, self.__component, True)
        return ShapeTextCursor(owner=self.__owner, component=xtext.createTextCursor())

    def get_glue_points(self) -> Tuple[GluePoint2, ...]:
        """
        Gets Glue Points.

        Raises:
            DrawError: If error occurs.

        Returns:
            Tuple[GluePoint2, ...]: Glue Points.

        Note:
            If a glue point can not be accessed then it is ignored.
        """
        return mDraw.Draw.get_glue_points(self.__component)

    def get_position_mm(self) -> Point:
        """
        Gets position in mm units

        Raises:
            PointError: If error occurs.

        Returns:
            Point: Position as Point in mm units
        """
        return mDraw.Draw.get_position(self.__component)

    def get_rotation(self) -> Angle:
        """
        Gets the rotation of a shape

        Args:
            shape (XShape): Shape

        Raises:
            ShapeError: If error occurs.

        Returns:
            Angle: Rotation angle.
        """
        return mDraw.Draw.get_rotation(self.__component)

    def get_shape_text(self) -> str:
        """
        Gets the text from inside a shape.

        Raises:
            DrawError: If error occurs getting shape text.

        Returns:
            str: Shape text
        """
        return mDraw.Draw.get_shape_text(shape=self.__component)

    def get_size_mm(self) -> Size:
        """
        Gets Size in mm units.

        Raises:
            SizeError: If error occurs.

        Returns:
            ~ooodev.utils.data_type.size.Size: Size in mm units
        """
        return mDraw.Draw.get_size(self.__component)

    def get_text_properties(self) -> XPropertySet:
        """
        Gets the properties associated with the text area inside the shape.

        Raises:
            PropertySetError: If error occurs.

        Returns:
            XPropertySet: Property Set
        """
        return mDraw.Draw.get_text_properties(self.__component)

    def get_transformation(self) -> HomogenMatrix3:
        """
        Gets a transformation matrix which seems to represent a clockwise rotation.

        Homogeneous matrix has three homogeneous lines

        Raises:
            ShapeError: If error occurs.

        Returns:
            HomogenMatrix3: Matrix
        """
        return mDraw.Draw.get_transformation(self.__component)

    def get_zorder(self) -> int:
        """
        Gets the z-order of a shape

        Raises:
            DrawError: If unable to get z-order.

        Returns:
            int: Z-Order
        """
        return mDraw.Draw.get_zorder(self.__component)

    def is_group(self) -> bool:
        """
        Gets if a shape is a Group Shape

        Returns:
            bool: ``True`` if shape is a group; Otherwise; ``False``.
        """
        return mDraw.Draw.is_group(self.__component)

    def is_image(self) -> bool:
        """
        Gets if a shape is an image (GraphicObjectShape).

        Returns:
            bool: ``True`` if shape is image; Otherwise, ``False``.
        """
        return mDraw.Draw.is_image(self.__component)

    def move_to_bottom(self) -> None:
        """
        Moves the shape to the bottom of the z-order

        Raises:
            ShapeMissingError: If unable to find shapes for slide.
            ShapeError: If any other error occurs.

        Returns:
            None:
        """
        if self.owner is None:
            raise mEx.ShapeError("Owner is None. Owner must be set before calling this method.")
        if mLo.Lo.is_uno_interfaces(self.owner.component, XDrawPage) is False:
            raise mEx.ShapeError(
                "Owner component is not a  is not a slide (XDrawPage). Owner must be a slide before calling this method."
            )
        mDraw.Draw.move_to_bottom(slide=self.owner.component, shape=self.__component)

    def move_to_top(self) -> None:
        """
        Moves the shape to the top of the z-order

        Raises:
            ShapeMissingError: If unable to find shapes for slide.
            ShapeError: If any other error occurs.

        Returns:
            None:
        """
        if self.owner is None:
            raise mEx.ShapeError("Owner is None. Owner must be set before calling this method.")
        if mLo.Lo.is_uno_interfaces(self.owner.component, XDrawPage) is False:
            raise mEx.ShapeError(
                "Owner component is not a  is not a slide (XDrawPage). Owner must be a slide before calling this method."
            )
        mDraw.Draw.move_to_top(slide=self.owner.component, shape=self.__component)

    def set_angle(self, angle: Angle | int) -> None:
        """
        Set the line style for a shape

        Args:
            angle (Angle | int): Angle to set.

        Raises:
            ShapeError: If error occurs.

        Returns:
            None:
        """
        mDraw.Draw.set_angle(shape=self.__component, angle=angle)

    def set_bitmap_color(self, name: DrawingBitmapKind | str) -> None:
        """
        Set bitmap color of a shape.

        Args:
            name (DrawingBitmapKind, str): Bitmap Name

        Raises:
            NameError: If ``name`` is not recognized.
            ShapeError: If any other error occurs.

        Returns:
            None:

        Note:
            Getting the bitmap color name can be a bit challenging.
            ``DrawingBitmapKind`` contains name displayed in the Bitmap color menu of Draw.

            The Easiest way to get the colors is to open Draw and see what bitmap color names are available
            on your system.
        """
        mDraw.Draw.set_bitmap_color(shape=self.__component, name=name)

    def set_bitmap_file_color(self, fnm: PathOrStr) -> None:
        """
        Set bitmap color from file.

        Args:
            fnm (PathOrStr): path to file.

        Raises:
            ShapeError: If error occurs.

        Returns:
            None:
        """
        mDraw.Draw.set_bitmap_file_color(shape=self.__component, fnm=fnm)

    def set_dashed_line(self, is_dashed: bool) -> None:
        """
        Set a dashed line

        Args:
            is_dashed (bool): Determines if line is to be dashed or solid.

        Raises:
            ShapeError: If error occurs.

        Returns:
            None:
        """
        mDraw.Draw.set_dashed_line(shape=self.__component, is_dashed=is_dashed)

    # region set_gradient_color()
    @overload
    def set_gradient_color(self, name: DrawingGradientKind | str) -> Gradient:
        """
        Set the gradient color of the shape.

        Args:
            name (DrawingGradientKind | str): Gradient color name.

        Returns:
            ~com.sun.star.awt.Gradient: Gradient instance that just had properties set.
        """
        ...

    @overload
    def set_gradient_color(self, start_color: mColor.Color, end_color: mColor.Color) -> Gradient:
        """
        Set the gradient color of the shape.

        Args:
            start_color (~ooodev.utils.color.Color): Start Color.
            end_color (~ooodev.utils.color.Color): End Color.

        Returns:
            ~com.sun.star.awt.Gradient: Gradient instance that just had properties set.
        """
        ...

    @overload
    def set_gradient_color(self, start_color: mColor.Color, end_color: mColor.Color, angle: Angle | int) -> Gradient:
        """
        Set the gradient color of the shape.

        Args:
            start_color (~ooodev.utils.color.Color): Start Color.
            end_color (~ooodev.utils.color.Color): End Color.
            angle (Angle, int): Gradient angle.

        Returns:
            ~com.sun.star.awt.Gradient: Gradient instance that just had properties set.
        """
        ...

    def set_gradient_color(self, *args, **kwargs) -> Gradient:
        """
        Set the gradient color of the shape.

        Args:
            name (DrawingGradientKind | str): Gradient color name.
            start_color (~ooodev.utils.color.Color): Start Color.
            end_color (~ooodev.utils.color.Color): End Color.
            angle (Angle, int): Gradient angle.

        Raises:
            NameError: If ``name`` is not recognized.
            ShapeError: If any other error occurs.

        Returns:
            ~com.sun.star.awt.Gradient: Gradient instance that just had properties set.

        Note:
            When using Gradient Name.

            Getting the gradient color name can be a bit challenging.
            ``DrawingGradientKind`` contains name displayed in the Gradient color menu of Draw.

            The Easiest way to get the colors is to open Draw and see what gradient color names are available
            on your system.

        See Also:
            :py:meth:`~.Draw.set_gradient_properties`
        """
        return mDraw.Draw.set_gradient_color(self.__component, *args, **kwargs)

    # endregion set_gradient_color()

    def set_gradient_properties(self, grad: Gradient) -> None:
        """
        Sets shapes gradient properties.

        Args:
            grad (~com.sun.star.awt.Gradient): Gradient properties to set.

        Returns:
            None:

        See Also:
            :py:meth:`~.Draw.set_gradient_color`
        """
        mDraw.Draw.set_gradient_properties(shape=self.__component, grad=grad)

    def set_hatch_color(self, name: DrawingHatchingKind | str) -> None:
        """
        Set hatching color of a shape.

        Args:
            name (DrawingHatchingKind, str): Hatching Name.

        Raises:
            NameError: If ``name`` is not recognized.
            ShapeError: If any other error occurs.

        Returns:
            None:

        Note:
            Getting the hatching color name can be a bit challenging.
            ``DrawingHatchingKind`` contains name displayed in the Hatching color menu of Draw.

            The Easiest way to get the colors is to open Draw and see what gradient color names are available
            on your system.
        """
        mDraw.Draw.set_hatch_color(shape=self.__component, name=name)

    def set_image(self, fnm: PathOrStr) -> None:
        """
        Sets the image of a shape.

        Args:
            fnm (PathOrStr): Path to image.

        Raises:
            ShapeError: If error occurs.

        Returns:
            None:
        """
        mDraw.Draw.set_image(shape=self.__component, fnm=fnm)

    def set_image_graphic(self, graphic: XGraphic) -> None:
        """
        Sets the image of a shape.

        Args:
            graphic (XGraphic): Graphic.

        Raises:
            ShapeError: If error occurs.

        Returns:
            None:
        """
        mDraw.Draw.set_image_graphic(shape=self.__component, graphic=graphic)

    def set_line_style(self, style: LineStyle) -> None:
        """
        Set the line style for a shape.

        Args:
            style (LineStyle): Line Style.

        Raises:
            ShapeError: If error occurs.

        Returns:
            None:
        """
        mDraw.Draw.set_line_style(shape=self.__component, style=style)

    # region set_position()

    @overload
    def set_position(self, pt: Point) -> None:
        """
        Sets Position of shape.

        Args:
            pt (point): Point that contains x and y positions in mm units.

        Returns:
            None:
        """
        ...

    @overload
    def set_position(self, x: int | UnitT, y: int | UnitT) -> None:
        """
        Sets Position of shape.

        Args:
            x (int, UnitT): X position in mm units or UnitT.
            y (int, UnitT): Y Position in mm units or UnitT.

        Returns:
            None:
        """
        ...

    def set_position(self, *args, **kwargs) -> None:
        """
        Sets Position of shape.

        Args:
            pt (point): Point that contains x and y positions in mm units.
            x (int, UnitT): X position in mm units or UnitT.
            y (int, UnitT): Y Position in mm units or UnitT.

        Raises:
            ShapeError: If error occurs.

        Returns:
            None:
        """
        mDraw.Draw.set_position(self.__component, *args, **kwargs)

    # endregion set_position()
    def set_props(self, **props) -> None:
        """
        Sets properties on a shape.

        Args:
            props (Any): Key value pairs of property name and property value

        Raises:
            MissingInterfaceError: if obj does not implement XPropertySet interface
            MultiError: If unable to set a property

        Returns:
            None:

        Example:
            .. code-block:: python

                set_props(Loop=True, MediaURL=FileIO.fnm_to_url(fnm))
        """
        mDraw.Draw.set_shape_props(self.__component, **props)

    def set_rotation(self, angle: Angle | int) -> None:
        """
        Set the rotation of a shape.

        Args:
            angle (Angle | int): Angle or int. An angle value from ``0`` to ``359``.

        Raises:
            ShapeError: If error occurs.

        Returns:
            None:
        """
        mDraw.Draw.set_rotation(shape=self.__component, angle=angle)

    # region Properties

    # region set_size()
    @overload
    def set_size(self, sz: SizeObj) -> None:
        ...

    @overload
    def set_size(self, width: int | UnitT, height: int | UnitT) -> None:
        ...

    def set_size(self, *args, **kwargs) -> None:
        """
        Sets set_size of shape.

        Args:
            sz (~ooodev.utils.data_type.size.Size): Size that contains width and height positions in mm units.
            width (int, UnitT): Width position in mm units or UnitT.
            height (int, UnitT): Height position in mm units or UnitT.

        Raises:
            ShapeError: If error occurs.

        Returns:
            None:
        """
        mDraw.Draw.set_size(self.__component, *args, **kwargs)

    # endregion set_size()

    def set_style(self, graphic_styles: XNameContainer, style_name: GraphicStyleKind | str) -> None:
        """
        Set the graphic style for a shape.

        Args:
            graphic_styles (XNameContainer): Graphic styles.
            style_name (GraphicStyleKind | str): Graphic Style Name.

        Raises:
            DrawError: If error occurs.

        Returns:
            None:
        """
        mDraw.Draw.set_style(shape=self.__component, graphic_styles=graphic_styles, style_name=style_name)

    def set_transparency(self, level: Intensity | int) -> None:
        """
        Sets the transparency level for the shape.
        Higher level means more transparent.

        Args:
            level (Intensity, int): Transparency value. Represents a intensity value from ``0`` to ``100``.

        Raises:
            ShapeError: If error occurs.

        Returns:
            None:
        """
        mDraw.Draw.set_transparency(shape=self.__component, level=level)

    def set_visible(self, is_visible: bool) -> None:
        """
        Set the line style for a shape.

        Args:
            is_visible (bool): Set is shape is visible or not.

        Raises:
            ShapeError: If error occurs.

        Returns:
            None:
        """
        mDraw.Draw.set_visible(shape=self.__component, is_visible=is_visible)

    def set_zorder(self, order: int) -> None:
        """
        Sets the z-order of a shape.

        Args:
            order (int): Z-Order.

        Raises:
            DrawError: If unable to set z-order.

        Returns:
            None:
        """
        mDraw.Draw.set_zorder(shape=self.__component, order=order)

    @property
    def owner(self) -> _T:
        """Component Owner"""
        return self.__owner

    # endregion Properties
