# region Import
from __future__ import annotations
from typing import Tuple, Type, cast, overload, TypeVar

from ooo.dyn.awt.point import Point

from ooodev.exceptions import ex as mEx
from ooodev.utils import props as mProps
from ooodev.units import UnitObj
from ooodev.units import UnitMM
from ooodev.units import UnitConvert
from ooodev.format.inner.kind.format_kind import FormatKind
from ...common.props.struct_size_props import StructSizeProps
from .struct_base import StructBase

# endregion Import

_TPointStruct = TypeVar(name="_TPointStruct", bound="PointStruct")


class PointStruct(StructBase):
    """
    Point struct.

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together Border Table properties.
    """

    # region init

    def __init__(self, x: int = 0, y: int = 0) -> None:
        """
        Constructor

        Args:
            x (int, optional): Specifies X coordinate. Default ``0``.
            y (int, optional): Specifies Y coordinate. Default ``0``.
        """
        super().__init__()
        self.prop_x = x
        self.prop_y = y

    # endregion init

    # region internal methods
    def _get_property_name(self) -> str:
        try:
            return self._property_name
        except AttributeError:
            self._property_name = "Position"
        return self._property_name

    # endregion internal methods

    # region dunder methods
    def __eq__(self, oth: object) -> bool:
        obj2 = None
        if isinstance(oth, PointStruct):
            obj2 = oth.get_uno_struct()
        if getattr(oth, "typeName", None) == "com.sun.star.awt.Point":
            obj2 = cast(PointStruct, oth)
        if obj2:
            obj1 = self.get_uno_struct()
            return obj1.X == obj2.X and obj1.Y == obj2.Y
        return NotImplemented

    # endregion dunder methods

    # region methods
    def get_uno_struct(self) -> Point:
        """
        Gets UNO ``Point`` from instance.

        Returns:
            Point: ``Point`` instance
        """
        x = self._get(self._props.width)
        y = self._get(self._props.height)
        inst = Point(X=x, Y=y)
        return inst

    # endregion methods

    # region overrides methods

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.drawing.Shape",)
        return self._supported_services_values

    # region apply()

    @overload
    def apply(self, obj: object) -> None:
        ...

    def apply(self, obj: object, **kwargs) -> None:
        """
        Applies tab properties to ``obj``

        Args:
            obj (object): UNO object.

        Returns:
            None:
        """
        name = self._get_property_name()
        if not name:
            return
        if not mProps.Props.has(obj, name):
            self._print_not_valid_srv("apply")
            return

        struct = self.get_uno_struct()
        props = {name: struct}
        super().apply(obj=obj, override_dv=props)

    # endregion apply()

    # endregion overrides methods

    # region static methods

    # region from_uno_struct()
    @overload
    @classmethod
    def from_uno_struct(cls: Type[_TPointStruct], value: Point) -> _TPointStruct:
        ...

    @overload
    @classmethod
    def from_uno_struct(cls: Type[_TPointStruct], value: Point, **kwargs) -> _TPointStruct:
        ...

    @classmethod
    def from_uno_struct(cls: Type[_TPointStruct], value: Point, **kwargs) -> _TPointStruct:
        """
        Converts a ``Point`` instance to a ``PointStruct``.

        Args:
            value (Point): UNO ``Point``.

        Returns:
            PointStruct: ``PointStruct`` set with ``Point`` properties.
        """
        inst = cls(**kwargs)
        inst._set(inst._props.height, value.Height)
        inst._set(inst._props.width, value.Width)
        return inst

    # endregion from_uno_struct()

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TPointStruct], obj: object) -> _TPointStruct:
        ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TPointStruct], obj: object, **kwargs) -> _TPointStruct:
        ...

    @classmethod
    def from_obj(cls: Type[_TPointStruct], obj: object, **kwargs) -> _TPointStruct:
        """
        Gets instance from object

        Args:
            obj (object): UNO object

        Raises:
            PropertyNotFoundError: If ``obj`` does not have required property

        Returns:
            PointStruct: ``PointStruct`` instance that represents ``obj`` Point properties.
        """
        # this nu is only used to get Property Name
        nu = cls(**kwargs)
        prop_name = nu._get_property_name()

        try:
            point = cast(PointStruct, mProps.Props.get(obj, prop_name))
        except mEx.PropertyNotFoundError:
            raise mEx.PropertyNotFoundError(prop_name, f"from_obj() obj as no {prop_name} property")

        return cls.from_uno_struct(point, **kwargs)

    # endregion from_obj()

    # endregion static methods

    # region Style methods
    def fmt_all(self: _TPointStruct, value: int) -> _TPointStruct:
        """
        Gets copy of instance with width and height set.

        Args:
            value (float, UnitObj): Specifies ``x`` and ``y`` values.

        Returns:
            PointStruct: Border Table
        """
        cp = self.copy()
        cp.prop_y = value
        cp.prop_x = value
        return cp

    def fmt_x(self: _TPointStruct, value: int) -> _TPointStruct:
        """
        Gets a copy of instance with height set.

        Args:
            value (float, UnitObj): Specifies ``x`` value.

        Returns:
            PointStruct:
        """
        cp = self.copy()
        cp.prop_x = value
        return cp

    def fmt_y(self: _TPointStruct, value: int) -> _TPointStruct:
        """
        Gets a copy of instance with width set.

        Args:
            value (float, UnitObj): Specifies ``y`` value.

        Returns:
            PointStruct:
        """
        cp = self.copy()
        cp.prop_y = value
        return cp

    # endregion Style methods

    # region Properties

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.STRUCT
        return self._format_kind_prop

    @property
    def prop_x(self) -> int:
        """Gets/Sets x value"""
        return self._get(self._props.width)

    @prop_x.setter
    def prop_x(self, value: int) -> None:
        self._set(self._props.width, value)

    @property
    def prop_y(self) -> int:
        """Gets/Sets y value."""
        return self._get(self._props.height)

    @prop_y.setter
    def prop_y(self, value: int) -> None:
        self._set(self._props.height, value)

    @property
    def _props(self) -> StructSizeProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = StructSizeProps(width="X", height="Y")
        return self._props_internal_attributes

    # endregion Properties
