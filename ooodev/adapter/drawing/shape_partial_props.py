"""
This class is a partial class for ``com.sun.star.drawing.Shape`` service.
It may be used in classes that implement the ``ooodev.adapter.drawing.shape_comp.ShapeComp`` class.
These properties are optional for a shape service.
"""
from __future__ import annotations
from typing import TYPE_CHECKING, Tuple
import uno


if TYPE_CHECKING:
    from com.sun.star.drawing import Shape  # service
    from com.sun.star.beans import PropertyValue
    from com.sun.star.container import XNameContainer
    from com.sun.star.style import XStyle
    from com.sun.star.drawing import HomogenMatrix3


class ShapePartialProps:
    """
    Partial Class for Shape Service.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: Shape) -> None:
        """
        Constructor

        Args:
            component (Shape): UNO Component that implements ``com.sun.star.drawing.Shape`` service.
        """
        self.__component = component

    # region Properties

    @property
    def interop_grab_bag(self) -> Tuple[PropertyValue, ...]:
        """
        Gets/Sets grab bag of shape properties, used as a string-any map for interim interop purposes.
        This property is intentionally not handled by the ODF filter. Any member that should be handled there should be
        first moved out from this grab bag to a separate property.
        """
        return self.__component.InteropGrabBag

    @interop_grab_bag.setter
    def interop_grab_bag(self, value: Tuple[PropertyValue, ...]) -> None:
        self.__component.InteropGrabBag = value

    @property
    def hyperlink(self) -> str:
        """
        Gets/Sets, this property lets you get and set a hyperlink for this shape.
        """
        return self.__component.Hyperlink

    @hyperlink.setter
    def hyperlink(self, value: str) -> None:
        self.__component.Hyperlink = value

    @property
    def layer_id(self) -> int:
        """
        Gets/Sets the ID of the Layer to which this Shape is attached.
        """
        return self.__component.LayerID

    @layer_id.setter
    def layer_id(self, value: int) -> None:
        self.__component.LayerID = value

    @property
    def layer_name(self) -> str:
        """
        Gets/Sets the name of the Layer to which this Shape is attached.
        """
        return self.__component.LayerName

    @layer_name.setter
    def layer_name(self, value: str) -> None:
        self.__component.LayerName = value

    @property
    def move_protect(self) -> bool:
        """
        Gets/Sets, With this set to ``True``, this Shape cannot be moved interactively in the user interface.
        """
        return self.__component.MoveProtect

    @move_protect.setter
    def move_protect(self, value: bool) -> None:
        self.__component.MoveProtect = value

    @property
    def name(self) -> str:
        """
        Gets/Sets the name of this Shape.
        """
        return self.__component.Name

    @name.setter
    def name(self, value: str) -> None:
        self.__component.Name = value

    @property
    def navigation_order(self) -> int:
        """
        Gets/Sets, this property stores the navigation order of this shape.
        If this value is negative, the navigation order for this shapes page is equal to the z-order.
        """
        return self.__component.NavigationOrder

    @navigation_order.setter
    def navigation_order(self, value: int) -> None:
        self.__component.NavigationOrder = value

    @property
    def printable(self) -> bool:
        """
        Gets/Sets, If this is ``False``, the Shape is not visible on printer outputs.
        """
        return self.__component.Printable

    @printable.setter
    def printable(self, value: bool) -> None:
        self.__component.Printable = value

    @property
    def relative_height(self) -> int:
        """
        Gets/Sets the relative height of the object.

        It is only valid if it is greater than zero.
        """
        return self.__component.RelativeHeight

    @relative_height.setter
    def relative_height(self, value: int) -> None:
        self.__component.RelativeHeight = value

    @property
    def relative_height_relation(self) -> int:
        """
        Gets/Sets the relation of the relative height of the object.

        It is only valid if RelativeHeight is greater than zero.
        """
        return self.__component.RelativeHeightRelation

    @relative_height_relation.setter
    def relative_height_relation(self, value: int) -> None:
        self.__component.RelativeHeightRelation = value

    @property
    def relative_width(self) -> int:
        """
        Gets/Sets the relative width of the object.

        It is only valid if it is greater than zero.
        """
        return self.__component.RelativeWidth

    @relative_width.setter
    def relative_width(self, value: int) -> None:
        self.__component.RelativeWidth = value

    @property
    def relative_width_relation(self) -> int:
        """
        Gets/Sets the relation of the relative width of the object.

        It is only valid if RelativeWidth is greater than zero.
        """
        return self.__component.RelativeWidthRelation

    @relative_width_relation.setter
    def relative_width_relation(self, value: int) -> None:
        self.__component.RelativeWidthRelation = value

    @property
    def shape_user_defined_attributes(self) -> XNameContainer:
        """
        Gets/Sets, this property stores xml attributes.

        They will be saved to and restored from automatic styles inside xml files.
        """
        return self.__component.ShapeUserDefinedAttributes

    @shape_user_defined_attributes.setter
    def shape_user_defined_attributes(self, value: XNameContainer) -> None:
        self.__component.ShapeUserDefinedAttributes = value

    @property
    def size_protect(self) -> bool:
        """
        Gets/Sets, With this set to ``True``, this Shape may not be sized interactively in the user interface.
        """
        return self.__component.SizeProtect

    @size_protect.setter
    def size_protect(self, value: bool) -> None:
        self.__component.SizeProtect = value

    @property
    def style(self) -> XStyle:
        """
        Gets/Sets, this property lets you get and set a style for this shape.
        """
        return self.__component.Style

    @style.setter
    def style(self, value: XStyle) -> None:
        self.__component.Style = value

    @property
    def transformation(self) -> HomogenMatrix3:
        """
        Gets/Sets, this property lets you get and set the transformation matrix for this shape.

        The transformation is a 3x3 homogeneous matrix and can contain translation, rotation, shearing and scaling.
        """
        return self.__component.Transformation

    @transformation.setter
    def transformation(self, value: HomogenMatrix3) -> None:
        self.__component.Transformation = value

    @property
    def visible(self) -> bool:
        """
        Gets/Sets, If this is ``False``, the Shape is not visible on screen outputs.

        Please note that the Shape may still be visible when printed, see Printable.
        """
        return self.__component.Visible

    @visible.setter
    def visible(self, value: bool) -> None:
        self.__component.Visible = value

    @property
    def z_order(self) -> int:
        """
        Gets/Sets the Zorder of this Shape.
        """
        return self.__component.ZOrder

    @z_order.setter
    def z_order(self, value: int) -> None:
        self.__component.ZOrder = value

    # endregion Properties
