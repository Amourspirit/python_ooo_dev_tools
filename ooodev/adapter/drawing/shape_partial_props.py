"""
This class is a partial class for ``com.sun.star.drawing.Shape`` service.
It may be used in classes that implement the ``ooodev.adapter.drawing.shape_comp.ShapeComp`` class.
These properties are optional for a shape service.
"""

from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING, Tuple
import contextlib

import uno
from ooodev.adapter.style.style_comp import StyleComp
from ooodev.adapter.drawing.homogen_matrix3_struct_comp import HomogenMatrix3StructComp
from ooodev.utils import info as mInfo
from ooodev.events.events import Events
from ooodev.utils import info as mInfo

if TYPE_CHECKING:
    from com.sun.star.drawing import Shape  # service
    from com.sun.star.beans import PropertyValue
    from com.sun.star.container import XNameContainer
    from com.sun.star.style import XStyle
    from com.sun.star.drawing import HomogenMatrix3

    from ooodev.events.args.key_val_args import KeyValArgs


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
        self.__event_provider = Events(self)
        self.__props = {}

        def on_comp_struct_changed(src: Any, event_args: KeyValArgs) -> None:
            prop_name = str(event_args.event_data["prop_name"])
            if hasattr(self.__component, prop_name):
                setattr(self.__component, prop_name, event_args.source.component)

        self.__fn_on_comp_struct_changed = on_comp_struct_changed
        # pylint: disable=no-member
        self.__event_provider.subscribe_event(
            "com_sun_star_drawing_HomogenMatrixLine3_changed", self.__fn_on_comp_struct_changed
        )

    # region Properties

    @property
    def interop_grab_bag(self) -> Tuple[PropertyValue, ...] | None:
        """
        Gets/Sets grab bag of shape properties, used as a string-any map for interim interop purposes.
        This property is intentionally not handled by the ODF filter. Any member that should be handled there should be
        first moved out from this grab bag to a separate property.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.InteropGrabBag
        return None

    @interop_grab_bag.setter
    def interop_grab_bag(self, value: Tuple[PropertyValue, ...]) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.InteropGrabBag = value

    @property
    def hyperlink(self) -> str | None:
        """
        Gets/Sets, this property lets you get and set a hyperlink for this shape.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.Hyperlink
        return None

    @hyperlink.setter
    def hyperlink(self, value: str) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.Hyperlink = value

    @property
    def layer_id(self) -> int | None:
        """
        Gets/Sets the ID of the Layer to which this Shape is attached.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.LayerID
        return None

    @layer_id.setter
    def layer_id(self, value: int) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.LayerID = value

    @property
    def layer_name(self) -> str | None:
        """
        Gets/Sets the name of the Layer to which this Shape is attached.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.LayerName
        return None

    @layer_name.setter
    def layer_name(self, value: str) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.LayerName = value

    @property
    def move_protect(self) -> bool | None:
        """
        Gets/Sets, With this set to ``True``, this Shape cannot be moved interactively in the user interface.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.MoveProtect
        return None

    @move_protect.setter
    def move_protect(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.MoveProtect = value

    @property
    def name(self) -> str | None:
        """
        Gets/Sets the name of this Shape.
        """
        with contextlib.suppress(AttributeError):
            return self.__component.Name
        return None

    @name.setter
    def name(self, value: str) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.Name = value

    @property
    def navigation_order(self) -> int | None:
        """
        Gets/Sets, this property stores the navigation order of this shape.
        If this value is negative, the navigation order for this shapes page is equal to the z-order.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.NavigationOrder
        return None

    @navigation_order.setter
    def navigation_order(self, value: int) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.NavigationOrder = value

    @property
    def printable(self) -> bool | None:
        """
        Gets/Sets, If this is ``False``, the Shape is not visible on printer outputs.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.Printable
        return None

    @printable.setter
    def printable(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.Printable = value

    @property
    def relative_height(self) -> int | None:
        """
        Gets/Sets the relative height of the object.

        It is only valid if it is greater than zero.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.RelativeHeight
        return None

    @relative_height.setter
    def relative_height(self, value: int) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.RelativeHeight = value

    @property
    def relative_height_relation(self) -> int | None:
        """
        Gets/Sets the relation of the relative height of the object.

        It is only valid if RelativeHeight is greater than zero.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.RelativeHeightRelation
        return None

    @relative_height_relation.setter
    def relative_height_relation(self, value: int) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.RelativeHeightRelation = value

    @property
    def relative_width(self) -> int | None:
        """
        Gets/Sets the relative width of the object.

        It is only valid if it is greater than zero.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.RelativeWidth
        return None

    @relative_width.setter
    def relative_width(self, value: int) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.RelativeWidth = value

    @property
    def relative_width_relation(self) -> int | None:
        """
        Gets/Sets the relation of the relative width of the object.

        It is only valid if RelativeWidth is greater than zero.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.RelativeWidthRelation
        return None

    @relative_width_relation.setter
    def relative_width_relation(self, value: int) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.RelativeWidthRelation = value

    @property
    def shape_user_defined_attributes(self) -> XNameContainer | None:
        """
        Gets/Sets, this property stores xml attributes.

        They will be saved to and restored from automatic styles inside xml files.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.ShapeUserDefinedAttributes
        return None

    @shape_user_defined_attributes.setter
    def shape_user_defined_attributes(self, value: XNameContainer) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.ShapeUserDefinedAttributes = value

    @property
    def size_protect(self) -> bool | None:
        """
        Gets/Sets, With this set to ``True``, this Shape may not be sized interactively in the user interface.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.SizeProtect
        return None

    @size_protect.setter
    def size_protect(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.SizeProtect = value

    @property
    def style(self) -> StyleComp | None:
        """
        Gets/Sets, this property lets you get and set a style for this shape.

        When setting property can be a ``StyleComp`` or ``XStyle``.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return StyleComp(self.__component.Style)
        return None

    @style.setter
    def style(self, value: XStyle | StyleComp) -> None:
        with contextlib.suppress(AttributeError):
            if mInfo.Info.is_instance(value, StyleComp):
                comp = value.component
            else:
                comp = cast("XStyle", value)
            self.__component.Style = comp

    @property
    def transformation(self) -> HomogenMatrix3StructComp | None:
        """
        Gets/Sets, this property lets you get and set the transformation matrix for this shape.

        The transformation is a 3x3 homogeneous matrix and can contain translation, rotation, shearing and scaling.

        When setting property can be a ``HomogenMatrix3`` or ``HomogenMatrix3StructComp``.

        **optional**

        Returns:
            HomogenMatrix3StructComp | None: The transformation matrix for this shape or None if property not supported.

        Hint:
            - ``HomogenMatrix3`` can be imported from ``ooo.dyn.drawing.homogen_matrix3``
        """
        key = "Transformation"
        if not hasattr(self.__component, key):
            return None
        prop = self.__props.get(key, None)
        if prop is None:
            prop = HomogenMatrix3StructComp(self.__component.Transformation, key, self.__event_provider)
            self.__props[key] = prop
        return cast(HomogenMatrix3StructComp, prop)

    @transformation.setter
    def transformation(self, value: HomogenMatrix3 | HomogenMatrix3StructComp) -> None:
        key = "Transformation"
        if not hasattr(self.__component, key):
            return None
        if mInfo.Info.is_instance(value, HomogenMatrix3StructComp):
            self.__component.Transformation = value.copy()
        else:
            self.__component.Transformation = cast("HomogenMatrix3", value)
        if key in self.__props:
            del self.__props[key]

    @property
    def visible(self) -> bool | None:
        """
        Gets/Sets, If this is ``False``, the Shape is not visible on screen outputs.

        Please note that the Shape may still be visible when printed, see Printable.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.Visible
        return None

    @visible.setter
    def visible(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.Visible = value

    @property
    def z_order(self) -> int | None:
        """
        Gets/Sets the Zorder of this Shape.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.ZOrder
        return None

    @z_order.setter
    def z_order(self, value: int) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.ZOrder = value

    # endregion Properties
