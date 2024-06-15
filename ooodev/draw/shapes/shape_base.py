from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING, TypeVar, Generic, overload, Tuple
import uno
from com.sun.star.drawing import XDrawPage
from com.sun.star.text import XText
from com.sun.star.drawing import XShape


from ooodev.draw.shapes.partial.export_jpg_partial import ExportJpgPartial
from ooodev.draw.shapes.partial.export_png_partial import ExportPngPartial
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.utils import info as mInfo
from ooodev.mock import mock_g
from ooodev.office import draw as mDraw
from ooodev.office.partial.office_document_prop_partial import OfficeDocumentPropPartial
from ooodev.units.angle import Angle
from ooodev.units.unit_mm import UnitMM
from ooodev.utils import gen_util as gUtil
from ooodev.utils.kind.drawing_bitmap_kind import DrawingBitmapKind
from ooodev.utils.kind.drawing_gradient_kind import DrawingGradientKind
from ooodev.utils.kind.drawing_hatching_kind import DrawingHatchingKind
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.partial.service_partial import ServicePartial
from ooodev.utils.partial.the_dictionary_partial import TheDictionaryPartial
from ooodev.adapter.awt.size_struct_generic_comp import SizeStructGenericComp
from ooodev.adapter.awt.point_struct_generic_comp import PointStructGenericComp

if TYPE_CHECKING:
    from com.sun.star.awt import Gradient
    from com.sun.star.awt import Point
    from com.sun.star.beans import XPropertySet
    from com.sun.star.container import XNameContainer
    from com.sun.star.drawing import GluePoint2
    from com.sun.star.drawing import HomogenMatrix3
    from com.sun.star.drawing import Shape  # Service
    from com.sun.star.graphic import XGraphic
    from com.sun.star.awt import Size as UnoSize
    from ooo.dyn.drawing.line_style import LineStyle
    from ooodev.draw.shapes.shape_text_cursor import ShapeTextCursor
    from ooodev.events.args.key_val_args import KeyValArgs
    from ooodev.events.lo_events import Events
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.proto.component_proto import ComponentT
    from ooodev.proto.size_obj import SizeObj
    from ooodev.units.unit_obj import UnitT
    from ooodev.utils import color as mColor
    from ooodev.utils.data_type.intensity import Intensity
    from ooodev.utils.data_type.size import Size
    from ooodev.utils.kind.graphic_style_kind import GraphicStyleKind
    from ooodev.utils.type_var import PathOrStr

_T = TypeVar("_T", bound="ComponentT")


class ShapeBase(
    LoInstPropsPartial,
    OfficeDocumentPropPartial,
    EventsPartial,
    ExportJpgPartial,
    ExportPngPartial,
    ServicePartial,
    TheDictionaryPartial,
    Generic[_T],
):
    def __init__(self, owner: _T, component: XShape, lo_inst: LoInst | None = None) -> None:
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo

        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        if not isinstance(owner, OfficeDocumentPropPartial):
            raise ValueError("owner must be an instance of OfficeDocumentPropPartial")
        OfficeDocumentPropPartial.__init__(self, owner.office_doc)
        EventsPartial.__init__(self)
        # pylint: disable=no-member
        events = cast("Events", self._EventsPartial__events)  # type: ignore
        ExportJpgPartial.__init__(self, component=component, events=events, lo_inst=self.lo_inst)
        ExportPngPartial.__init__(self, component=component, events=events, lo_inst=self.lo_inst)
        ServicePartial.__init__(self, component=component, lo_inst=self.lo_inst)
        TheDictionaryPartial.__init__(self)
        self.__owner = owner
        self.__component = component
        self.__props = {}
        self._apply_shape_name()

        def on_comp_struct_changed(src: Any, event_args: KeyValArgs) -> None:
            prop_name = str(event_args.event_data["prop_name"])
            setattr(self, prop_name, src.component)

        self.__fn_on_comp_struct_changed = on_comp_struct_changed
        # pylint: disable=no-member
        self.subscribe_event("generic_com_sun_star_awt_Size_changed", self.__fn_on_comp_struct_changed)
        self.subscribe_event("generic_com_sun_star_awt_Point_changed", self.__fn_on_comp_struct_changed)

    def __repr__(self) -> str:
        return f"{self.__class__.__name__}({self.shape_type})"

    def _clone(self) -> XShape:
        """
        Clones the shape.

        Does not apply to all shapes, just known shapes.
        """
        try:
            page = self.__owner.component
        except AttributeError as e:
            raise mEx.NotSupportedError("Owner must be a draw page") from e

        old_shape = self.__component
        pt = old_shape.getPosition()
        sz = old_shape.getSize()
        shape = self.lo_inst.create_instance_msf(XShape, old_shape.getShapeType(), raise_err=True)
        shape.setPosition(pt)
        shape.setSize(sz)
        try:
            page.add(shape)
        except Exception as exc:
            raise mEx.NotSupportedError("Owner must be a draw page") from exc
        return shape

    def _generate_shape_name(self) -> str:
        shape = self.__component
        # Get the shape name from the shape type.

        shape_type = shape.getShapeType()
        # only get that laster part of the name after the last dot.
        shape_name = shape_type.rsplit(".")[-1]
        return f"{shape_name}_{gUtil.Util.generate_random_string(10)}"

    def _apply_shape_name(self, keep_existing: bool = True) -> None:
        shape = cast("Shape", self.__component)
        if not shape.Name or not keep_existing:
            shape.Name = self._generate_shape_name()

    def get_lo_inst(self) -> LoInst:
        return self.lo_inst

    def get_owner(self) -> _T:
        return self.__owner

    def get_shape_text_cursor(self) -> ShapeTextCursor[_T]:
        """
        Gets a cursor object for this text.

        Returns:
            ShapeTextCursor: Cursor.
        """
        # pylint: disable=import-outside-toplevel
        # pylint: disable=redefined-outer-name
        # avoid a circular import
        from ooodev.draw.shapes.shape_text_cursor import ShapeTextCursor

        lo = self.get_lo_inst()
        xtext = lo.qi(XText, self.__component, True)
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
        if self.lo_inst.is_uno_interfaces(self.owner.component, XDrawPage) is False:
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
        if self.lo_inst.is_uno_interfaces(self.owner.component, XDrawPage) is False:
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

    # region set_size()
    @overload
    def set_size(self, sz: SizeObj) -> None:
        """
        Sets set_size of shape.

        Args:
            sz (~ooodev.utils.data_type.size.Size): Size that contains width and height positions in mm units.

        Raises:
            ShapeError: If error occurs.

        Returns:
            None:
        """
        ...

    @overload
    def set_size(self, width: int | UnitT, height: int | UnitT) -> None:
        """
        Sets set_size of shape.

        Args:
            width (int, UnitT): Width position in mm units or UnitT.
            height (int, UnitT): Height position in mm units or UnitT.

        Raises:
            ShapeError: If error occurs.

        Returns:
            None:
        """
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

    def get_shape_type(self) -> str:
        """Get the shape type. This is usually a service name and is manually set by the class."""
        raise NotImplementedError

    # region Properties

    @property
    def owner(self) -> _T:
        """Component Owner"""
        return self.__owner

    # @property
    # def size(self) -> GenericUnitSize[UnitMM, float]:
    #     """Gets the size of the shape in ``UnitMM`` Values."""
    #     sz = self.__component.getSize()
    #     return GenericUnitSize(UnitMM.from_mm100(sz.Width), UnitMM.from_mm100(sz.Height))

    @property
    def size(self) -> SizeStructGenericComp[UnitMM]:
        """
        Gets/Sets the size of the shape in ``UnitMM`` Values.

        When setting this value, it can be set with a ``com.sun.star.awt.Size`` instance or a ``SizeStructGenericComp`` instance.

        The Size can be set by just setting a Size property.

        Returns:
            SizeStructGenericComp[UnitMM]: Size in ``UnitMM`` Values.

        Example:

            Can be set by just setting a Size property using int or ``UnitMM``.

            .. code-block:: python

                shape.size.width = 1000 # 10 mm
                shape.size.height = UnitMM(20) # 20 mm

            Can also be set using any ``UnitT`` object.

            .. code-block:: python

                shape.size.width = UnitCM(1.2)
                shape.size.height = UnitMM(40)

            Can also be set using a ``com.sun.star.awt.Size`` struct.

            .. code-block:: python

                shape.size = Size(1000, 2000) # in 1/100mm
        """
        key = "size"
        sz = self.__component.getSize()
        prop = self.__props.get(key, None)
        if prop is None:
            prop = SizeStructGenericComp(sz, UnitMM, key, self)
            self.__props[key] = prop
        return cast(SizeStructGenericComp[UnitMM], prop)

    @size.setter
    def size(self, value: UnoSize | SizeStructGenericComp) -> None:
        key = "size"
        if mInfo.Info.is_instance(value, SizeStructGenericComp):
            self.__component.setSize(value.copy())
        else:
            self.__component.setSize(cast("UnoSize", value))
        if key in self.__props:
            del self.__props[key]

    @property
    def position(self) -> PointStructGenericComp[UnitMM]:
        """
        Gets/Sets the position of the shape in ``UnitMM`` Values.

        When setting this value, it can be set with a ``com.sun.star.awt.Position`` instance or a ``PointStructGenericComp`` instance.

        The Position can be set by just setting a Position property.

        Returns:
            PointStructGenericComp[UnitMM]: Position in ``UnitMM`` Values.

        Note:
            This is the position as reported by the shape. This is not the same as the position of the shape on the page.
            The position of the shape on the page includes is from the page margins.

            For instance if the page has a margin of 10mm and the shape is at position (15, 15) then the position of the shape
            in the Draw Dialog box is (5, 5) where as this position property will report (15, 15).

        Example:

            Can be set by just setting a Position property using int or ``UnitMM``.

            .. code-block:: python

                shape.position.x = 1000 # 10 mm
                shape.position.y = UnitMM(20) # 20 mm

            Can also be set using any ``UnitT`` object.

            .. code-block:: python

                shape.position.x = UnitCM(1.2)
                shape.position.y = UnitMM(40)

            Can also be set using a ``com.sun.star.awt.Position`` struct.

            .. code-block:: python

                shape.position = Position(1000, 2000) # in 1/100mm
        """
        key = "position"
        pos = self.__component.getPosition()
        prop = self.__props.get(key, None)
        if prop is None:
            prop = PointStructGenericComp(pos, UnitMM, key, self)
            self.__props[key] = prop
        return cast(PointStructGenericComp[UnitMM], prop)

    @position.setter
    def position(self, value: Point | PointStructGenericComp) -> None:
        key = "position"
        if mInfo.Info.is_instance(value, PointStructGenericComp):
            self.__component.setPosition(value.copy())
        else:
            self.__component.setPosition(cast("Point", value))
        if key in self.__props:
            del self.__props[key]

    @property
    def shape_type(self) -> str:
        """Gets the shape type."""
        return self.__component.ShapeType  # type: ignore

    @property
    def component(self) -> XShape:
        """Gets the component."""
        return self.__component

    # endregion Properties


if mock_g.FULL_IMPORT:
    from ooodev.draw.shapes.shape_text_cursor import ShapeTextCursor
