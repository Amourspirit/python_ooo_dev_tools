"""
Module for Image Crop (``GraphicCrop``) struct

.. versionadded:: 0.9.0
"""

# region imports
from __future__ import annotations
from typing import Any, Tuple, Type, cast, overload, TypeVar, TYPE_CHECKING

from ooo.dyn.text.graphic_crop import GraphicCrop

from ooodev.exceptions import ex as mEx
from ooodev.format.inner.common.props.struct_crop_props import StructCropProps
from ooodev.format.inner.direct.structs.struct_base import StructBase
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.units.unit_convert import UnitConvert
from ooodev.units.unit_mm import UnitMM
from ooodev.utils import props as mProps

if TYPE_CHECKING:
    from ooodev.units.unit_obj import UnitT

# endregion imports

_TCropStruct = TypeVar(name="_TCropStruct", bound="CropStruct")


class CropStruct(StructBase):
    """
    Crop struct.

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together Border Table properties.
    """

    # region init

    def __init__(
        self,
        *,
        left: float | UnitT = 0.0,
        right: float | UnitT = 0.0,
        top: float | UnitT = 0.0,
        bottom: float | UnitT = 0.0,
        all: float | UnitT | None = None,
    ) -> None:
        """
        Constructor

        Args:
            left (float, UnitT, optional): Specifies left crop in ``mm`` units or :ref:`proto_unit_obj`. Default ``0.0``.
            right (float, UnitT, optional): Specifies right crop in ``mm`` units or :ref:`proto_unit_obj`. Default ``0.0``.
            top (float, UnitT, optional): Specifies top crop in ``mm`` units or :ref:`proto_unit_obj`. Default ``0.0``.
            bottom (float, UnitT, optional): Specifies bottom crop in ``mm`` units or :ref:`proto_unit_obj`. Default ``0.0``.
            all (float, UnitT, optional): Specifies ``left``, ``right``, ``top``, and ``bottom`` in ``mm`` units or :ref:`proto_unit_obj`. If set all other parameters are ignored.
        """
        super().__init__()
        if all is not None:
            self.prop_left = all
            self.prop_right = all
            self.prop_top = all
            self.prop_bottom = all
        else:
            self.prop_left = left
            self.prop_right = right
            self.prop_top = top
            self.prop_bottom = bottom

    # endregion init

    # region internal methods
    def _get_property_name(self) -> str:
        try:
            return self._property_name
        except AttributeError:
            self._property_name = "GraphicCrop"
        return self._property_name

    # endregion internal methods

    # region dunder methods
    def __eq__(self, oth: object) -> bool:
        obj2 = None
        if isinstance(oth, CropStruct):
            obj2 = oth.get_uno_struct()
        if getattr(oth, "typeName", None) == "com.sun.star.text.GraphicCrop":
            obj2 = cast(GraphicCrop, oth)
        if obj2:
            obj1 = self.get_uno_struct()
            return (
                obj1.Left == obj2.Left
                and obj1.Right == obj2.Right
                and obj1.Top == obj2.Top
                and obj1.Bottom == obj2.Bottom
            )
        return NotImplemented

    # endregion dunder methods

    # region methods
    def get_uno_struct(self) -> GraphicCrop:
        """
        Gets UNO ``GraphicCrop`` from instance.

        Returns:
            GraphicCrop: ``GraphicCrop`` instance
        """
        return GraphicCrop(
            Top=self._get(self._props.top),
            Bottom=self._get(self._props.bottom),
            Left=self._get(self._props.left),
            Right=self._get(self._props.right),
        )

    # endregion methods

    # region overrides methods

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.text.TextGraphicObject",)
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
        if not mProps.Props.has(obj, self._get_property_name()):
            self._print_not_valid_srv("apply")
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
    def from_uno_struct(cls: Type[_TCropStruct], value: GraphicCrop) -> _TCropStruct: ...

    @overload
    @classmethod
    def from_uno_struct(cls: Type[_TCropStruct], value: GraphicCrop, **kwargs) -> _TCropStruct: ...

    @classmethod
    def from_uno_struct(cls: Type[_TCropStruct], value: GraphicCrop, **kwargs) -> _TCropStruct:
        """
        Converts a ``GraphicCrop`` instance to a ``CropStruct``.

        Args:
            value (GraphicCrop): UNO ``GraphicCrop``.

        Returns:
            CropStruct: ``CropStruct`` set with ``GraphicCrop`` properties.
        """
        inst = cls(**kwargs)
        inst._set(inst._props.left, value.Left)
        inst._set(inst._props.right, value.Right)
        inst._set(inst._props.top, value.Top)
        inst._set(inst._props.bottom, value.Bottom)
        return inst

    # endregion from_uno_struct()

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TCropStruct], obj: Any) -> _TCropStruct: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TCropStruct], obj: Any, **kwargs) -> _TCropStruct: ...

    @classmethod
    def from_obj(cls: Type[_TCropStruct], obj: Any, **kwargs) -> _TCropStruct:
        """
        Gets instance from object

        Args:
            obj (object): UNO object

        Raises:
            PropertyNotFoundError: If ``obj`` does not have required property

        Returns:
            CropStruct: ``CropStruct`` instance that represents ``obj`` crop properties.
        """
        # this nu is only used to get Property Name
        nu = cls(**kwargs)
        prop_name = nu._get_property_name()

        try:
            grad = cast(GraphicCrop, mProps.Props.get(obj, prop_name))
        except mEx.PropertyNotFoundError as e:
            raise mEx.PropertyNotFoundError(prop_name, f"from_obj() obj as no {prop_name} property") from e

        return cls.from_uno_struct(grad, **kwargs)

    # endregion from_obj()

    # endregion static methods

    # region Style methods
    def fmt_all(self: _TCropStruct, value: float | UnitT) -> _TCropStruct:
        """
        Gets copy of instance with left, right, top, bottom set.

        Args:
            value (float, UnitT): Specifies crop in ``mm`` units or :ref:`proto_unit_obj`.

        Returns:
            CropStruct: Border Table
        """
        cp = self.copy()
        cp.prop_top = value
        cp.prop_bottom = value
        cp.prop_left = value
        cp.prop_right = value
        return cp

    def fmt_top(self: _TCropStruct, value: float | UnitT) -> _TCropStruct:
        """
        Gets a copy of instance with top set.

        Args:
            value (float, UnitT): Specifies top crop in ``mm`` units or :ref:`proto_unit_obj`.

        Returns:
            CropStruct:
        """
        cp = self.copy()
        cp.prop_top = value
        return cp

    def fmt_bottom(self: _TCropStruct, value: float | UnitT) -> _TCropStruct:
        """
        Gets a copy of instance with bottom set.

        Args:
            value (float, UnitT): Specifies bottom crop in ``mm`` units or :ref:`proto_unit_obj`.

        Returns:
            CropStruct:
        """
        cp = self.copy()
        cp.prop_bottom = value
        return cp

    def fmt_left(self: _TCropStruct, value: float | UnitT) -> _TCropStruct:
        """
        Gets a copy of instance with left set.

        Args:
            value (float, UnitT): Specifies left crop in ``mm`` units or :ref:`proto_unit_obj`.

        Returns:
            CropStruct:
        """
        cp = self.copy()
        cp.prop_left = value
        return cp

    def fmt_right(self: _TCropStruct, value: float | UnitT) -> _TCropStruct:
        """
        Gets a copy of instance with right set.

        Args:
            value (float, UnitT): Specifies right crop in ``mm`` units or :ref:`proto_unit_obj`.

        Returns:
            CropStruct:
        """
        cp = self.copy()
        cp.prop_right = value
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
    def prop_left(self) -> UnitMM:
        """Gets/Sets left value in ``mm`` units."""
        pv = self._get(self._props.left)
        return UnitMM.from_mm100(pv)

    @prop_left.setter
    def prop_left(self, value: float | UnitT) -> None:
        try:
            self._set(self._props.left, value.get_value_mm100())  # type: ignore
        except AttributeError:
            self._set(self._props.left, UnitConvert.convert_mm_mm100(value))  # type: ignore

    @property
    def prop_right(self) -> UnitMM:
        """Gets/Sets right value in ``mm`` units."""
        pv = self._get(self._props.right)
        return UnitMM.from_mm100(pv)

    @prop_right.setter
    def prop_right(self, value: float | UnitT) -> None:
        try:
            self._set(self._props.right, value.get_value_mm100())  # type: ignore
        except AttributeError:
            self._set(self._props.right, UnitConvert.convert_mm_mm100(value))  # type: ignore

    @property
    def prop_top(self) -> UnitMM:
        """Gets/Sets top value in ``mm`` units."""
        pv = self._get(self._props.top)
        return UnitMM.from_mm100(pv)

    @prop_top.setter
    def prop_top(self, value: float | UnitT) -> None:
        try:
            self._set(self._props.top, value.get_value_mm100())  # type: ignore
        except AttributeError:
            self._set(self._props.top, UnitConvert.convert_mm_mm100(value))  # type: ignore

    @property
    def prop_bottom(self) -> UnitMM:
        """Gets/Sets bottom value in ``mm`` units."""
        pv = self._get(self._props.bottom)
        return UnitMM.from_mm100(pv)

    @prop_bottom.setter
    def prop_bottom(self, value: float | UnitT) -> None:
        try:
            self._set(self._props.bottom, value.get_value_mm100())  # type: ignore
        except AttributeError:
            self._set(self._props.bottom, UnitConvert.convert_mm_mm100(value))  # type: ignore

    @property
    def _props(self) -> StructCropProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = StructCropProps(left="Left", top="Top", right="Right", bottom="Bottom")
        return self._props_internal_attributes

    # endregion Properties
