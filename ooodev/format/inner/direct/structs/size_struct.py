# region Import
from __future__ import annotations
from typing import Any, Tuple, Type, cast, overload, TypeVar

from ooo.dyn.awt.size import Size

from ooodev.exceptions import ex as mEx
from ooodev.utils import props as mProps
from ooodev.units.unit_obj import UnitT
from ooodev.units.unit_mm import UnitMM
from ooodev.units.unit_convert import UnitConvert
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.common.props.struct_size_props import StructSizeProps
from ooodev.format.inner.direct.structs.struct_base import StructBase

# endregion Import

_TSizeStruct = TypeVar("_TSizeStruct", bound="SizeStruct")


class SizeStruct(StructBase):
    """
    Size struct.

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together Border Table properties.
    """

    # region init

    def __init__(
        self,
        width: float | UnitT = 0.0,
        height: float | UnitT = 0.0,
        all: float | UnitT | None = None,
    ) -> None:
        """
        Constructor

        Args:
            width (float, UnitT, optional): Specifies width crop in ``mm`` units or :ref:`proto_unit_obj`. Default ``0.0``.
            height (float, UnitT, optional): Specifies height crop in ``mm`` units or :ref:`proto_unit_obj`. Default ``0.0``.
            all (float, UnitT, optional): Specifies ``width`` and ``height`` in ``mm`` units or :ref:`proto_unit_obj`. If set all other parameters are ignored.
        """
        super().__init__()
        if all is not None:
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
        return Size(
            Width=self._get(self._props.width),
            Height=self._get(self._props.height),
        )

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
    def apply(self, obj: Any) -> None:  # type: ignore
        ...

    def apply(self, obj: Any, **kwargs) -> None:
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

        grad = self.get_uno_struct()
        props = {name: grad}
        super().apply(obj=obj, override_dv=props)

    # endregion apply()

    # endregion overrides methods

    # region static methods

    # region from_uno_struct()
    @overload
    @classmethod
    def from_uno_struct(cls: Type[_TSizeStruct], value: Size) -> _TSizeStruct: ...

    @overload
    @classmethod
    def from_uno_struct(cls: Type[_TSizeStruct], value: Size, **kwargs) -> _TSizeStruct: ...

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
    def from_obj(cls: Type[_TSizeStruct], obj: Any) -> _TSizeStruct: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TSizeStruct], obj: Any, **kwargs) -> _TSizeStruct: ...

    @classmethod
    def from_obj(cls: Type[_TSizeStruct], obj: Any, **kwargs) -> _TSizeStruct:
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
            size = cast(Size, mProps.Props.get(obj, prop_name))
        except mEx.PropertyNotFoundError as e:
            raise mEx.PropertyNotFoundError(prop_name, f"from_obj() obj as no {prop_name} property") from e

        return cls.from_uno_struct(size, **kwargs)

    # endregion from_obj()

    # endregion static methods

    # region Style methods
    def fmt_all(self: _TSizeStruct, value: float | UnitT) -> _TSizeStruct:
        """
        Gets copy of instance with width and height set.

        Args:
            value (float, UnitT): Specifies crop in ``mm`` units or :ref:`proto_unit_obj`.

        Returns:
            SizeStruct: Border Table
        """
        cp = self.copy()
        cp.prop_width = value
        cp.prop_height = value
        return cp

    def fmt_height(self: _TSizeStruct, value: float | UnitT) -> _TSizeStruct:
        """
        Gets a copy of instance with height set.

        Args:
            value (float, UnitT): Specifies height in ``mm`` units or :ref:`proto_unit_obj`.

        Returns:
            SizeStruct:
        """
        cp = self.copy()
        cp.prop_height = value
        return cp

    def fmt_width(self: _TSizeStruct, value: float | UnitT) -> _TSizeStruct:
        """
        Gets a copy of instance with width set.

        Args:
            value (float, UnitT): Specifies width in ``mm`` units or :ref:`proto_unit_obj`.

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
    def prop_height(self, value: float | UnitT) -> None:
        try:
            self._set(self._props.height, value.get_value_mm100())  # type: ignore
        except AttributeError:
            self._set(self._props.height, UnitConvert.convert_mm_mm100(value))  # type: ignore

    @property
    def prop_width(self) -> UnitMM:
        """Gets/Sets width value in ``mm`` units."""
        pv = self._get(self._props.width)
        return UnitMM.from_mm100(pv)

    @prop_width.setter
    def prop_width(self, value: float | UnitT) -> None:
        try:
            self._set(self._props.width, value.get_value_mm100())  # type: ignore
        except AttributeError:
            self._set(self._props.width, UnitConvert.convert_mm_mm100(value))  # type: ignore

    @property
    def _props(self) -> StructSizeProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = StructSizeProps(width="Width", height="Height")
        return self._props_internal_attributes

    # endregion Properties
