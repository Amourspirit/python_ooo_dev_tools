"""
Module for Image Crop (``GraphicCrop``) struct

.. versionadded:: 0.9.0
"""
# region imports
from __future__ import annotations
from typing import Tuple, Type, cast, overload, TypeVar

import uno
from ooo.dyn.awt.size import Size

from ....exceptions import ex as mEx
from ....proto.unit_obj import UnitObj
from ....utils import props as mProps
from ....utils.data_type.unit_mm import UnitMM
from ....utils.unit_convert import UnitConvert
from ...kind.format_kind import FormatKind
from ..common.props.struct_size_props import StructSizeProps
from .struct_base import StructBase

# endregion imports

_TSizeStruct = TypeVar(name="_TSizeStruct", bound="SizeStruct")


class SizeStruct(StructBase):
    """
    Size struct.

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together Border Table properties.
    """

    # region init

    def __init__(
        self,
        width: float | UnitObj = 0.0,
        height: float | UnitObj = 0.0,
        all: float | UnitObj = None,
    ) -> None:
        """
        Constructor

        Args:
            width (float, UnitObj, optional): Specifies width crop in ``mm`` units or :ref:`proto_unit_obj`. Default ``0.0``.
            height (float, UnitObj, optional): Specifies height crop in ``mm`` units or :ref:`proto_unit_obj`. Default ``0.0``.
            all (float, UnitObj, optional): Specifies ``width`` and ``height`` in ``mm`` units or :ref:`proto_unit_obj`. If set all other parameters are ignored.
        """
        super().__init__()
        if not all is None:
            self.prop_left = all
            self.prop_right = all
        else:
            self.prop_left = width
            self.prop_right = height

    # endregion init

    # region internal methods
    def _get_property_name(self) -> str:
        try:
            return self._property_name
        except AttributeError:
            self._property_name = "Size"
        return self._property_name

    # endregion internal methods

    # region dunder methods
    def __eq__(self, oth: object) -> bool:
        obj2 = None
        if isinstance(oth, SizeStruct):
            obj2 = oth.get_uno_struct()
        if getattr(oth, "typeName", None) == "com.sun.star.awt.Size":
            obj2 = cast(Size, oth)
        if obj2:
            obj1 = self.get_uno_struct()
            return obj1.Width == obj2.Width and obj1.Height == obj2.Height
        return NotImplemented

    # endregion dunder methods

    # region methods
    def get_uno_struct(self) -> Size:
        """
        Gets UNO ``Size`` from instance.

        Returns:
            Size: ``Size`` instance
        """
        inst = Size(Width=self._get(self._props.width), Height=self._get(self._props.height))
        return inst

    # endregion methods

    # region overrides methods

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.text.TextGraphicObjects",)
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
        if not mProps.Props.has(obj, self._get_property_name()):
            self._print_not_valid_obj("apply")
            return

        grad = self.get_uno_struct()
        props = {self._get_property_name(): grad}
        super().apply(obj=obj, override_dv=props)

    # endregion apply()

    # endregion overrides methods

    # region static methods

    # region from_uno_struct()
    @overload
    @classmethod
    def from_uno_struct(cls: Type[_TSizeStruct], value: Size) -> _TSizeStruct:
        ...

    @overload
    @classmethod
    def from_uno_struct(cls: Type[_TSizeStruct], value: Size, **kwargs) -> _TSizeStruct:
        ...

    @classmethod
    def from_uno_struct(cls: Type[_TSizeStruct], value: Size, **kwargs) -> _TSizeStruct:
        """
        Converts a ``Size`` instance to a ``SizeStruct``.

        Args:
            value (Size): UNO ``Size``.

        Returns:
            SizeStruct: ``SizeStruct`` set with ``Size`` properties.
        """
        inst = cls(**kwargs)
        inst._set(inst._props.height, value.Height)
        inst._set(inst._props.width, value.Width)
        return inst

    # endregion from_uno_struct()

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TSizeStruct], obj: object) -> _TSizeStruct:
        ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TSizeStruct], obj: object, **kwargs) -> _TSizeStruct:
        ...

    @classmethod
    def from_obj(cls: Type[_TSizeStruct], obj: object, **kwargs) -> _TSizeStruct:
        """
        Gets instance from object

        Args:
            obj (object): UNO object

        Raises:
            PropertyNotFoundError: If ``obj`` does not have required property

        Returns:
            SizeStruct: ``SizeStruct`` instance that represents ``obj`` Size properties.
        """
        # this nu is only used to get Property Name
        nu = cls(**kwargs)
        prop_name = nu._get_property_name()

        try:
            size = cast(SizeStruct, mProps.Props.get(obj, prop_name))
        except mEx.PropertyNotFoundError:
            raise mEx.PropertyNotFoundError(prop_name, f"from_obj() obj as no {prop_name} property")

        return cls.from_uno_struct(size, **kwargs)

    # endregion from_obj()

    # endregion static methods

    # region Style methods
    def fmt_all(self: _TSizeStruct, value: float | UnitObj) -> _TSizeStruct:
        """
        Gets copy of instance with width and height set.

        Args:
            value (float, UnitObj): Specifies crop in ``mm`` units or :ref:`proto_unit_obj`.

        Returns:
            SizeStruct: Border Table
        """
        cp = self.copy()
        cp.prop_width = value
        cp.prop_height = value
        return cp

    def fmt_height(self: _TSizeStruct, value: float | UnitObj) -> _TSizeStruct:
        """
        Gets a copy of instance with height set.

        Args:
            value (float, UnitObj): Specifies height in ``mm`` units or :ref:`proto_unit_obj`.

        Returns:
            SizeStruct:
        """
        cp = self.copy()
        cp.prop_height = value
        return cp

    def fmt_width(self: _TSizeStruct, value: float | UnitObj) -> _TSizeStruct:
        """
        Gets a copy of instance with width set.

        Args:
            value (float, UnitObj): Specifies width in ``mm`` units or :ref:`proto_unit_obj`.

        Returns:
            SizeStruct:
        """
        cp = self.copy()
        cp.prop_width = value
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
    def prop_height(self) -> UnitMM:
        """Gets/Sets height value in ``mm`` units."""
        pv = self._get(self._props.height)
        return UnitMM.from_mm100(pv)

    @prop_height.setter
    def prop_height(self, value: float | UnitObj) -> None:
        try:
            self._set(self._props.height, value.get_value_mm100())
        except AttributeError:
            self._set(self._props.height, UnitConvert.convert_mm_mm100(value))

    @property
    def prop_width(self) -> UnitMM:
        """Gets/Sets width value in ``mm`` units."""
        pv = self._get(self._props.width)
        return UnitMM.from_mm100(pv)

    @prop_width.setter
    def prop_width(self, value: float | UnitObj) -> None:
        try:
            self._set(self._props.width, value.get_value_mm100())
        except AttributeError:
            self._set(self._props.width, UnitConvert.convert_mm_mm100(value))

    @property
    def _props(self) -> StructSizeProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = StructSizeProps(width="Width", height="Height")
        return self._props_internal_attributes

    # endregion Properties
