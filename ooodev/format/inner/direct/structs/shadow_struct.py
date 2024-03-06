# region Import
from __future__ import annotations
from typing import Any, Dict, Tuple, Type, cast, overload, TypeVar

import uno
from ooo.dyn.table.shadow_format import ShadowFormat
from ooo.dyn.table.shadow_location import ShadowLocation

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.utils import props as mProps
from ooodev.utils.color import Color, StandardColor
from ooodev.units.unit_obj import UnitT
from ooodev.units.unit_mm import UnitMM
from ooodev.units.unit_mm100 import UnitMM100
from ooodev.units.unit_convert import UnitConvert
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.format.inner.direct.structs.struct_base import StructBase

# endregion Import

_TShadowStruct = TypeVar(name="_TShadowStruct", bound="ShadowStruct")


class ShadowStruct(StructBase):
    """
    Shadow struct

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together properties.

    .. versionadded:: 0.9.0
    """

    # region init

    def __init__(
        self,
        *,
        location: ShadowLocation = ShadowLocation.BOTTOM_RIGHT,
        color: Color = StandardColor.GRAY,
        transparent: bool = False,
        width: float | UnitT = 1.76,
    ) -> None:
        """
        Constructor

        Args:
            location (ShadowLocation, optional): contains the location of the shadow.
                Default to ``ShadowLocation.BOTTOM_RIGHT``.
            color (:py:data:`~.utils.color.Color`, optional):contains the color value of the shadow. Defaults to ``StandardColor.GRAY``.
            transparent (bool, optional): Shadow transparency. Defaults to False.
            width (float, UnitT, optional): contains the size of the shadow (in ``mm`` units)
                or :ref:`proto_unit_obj`. Defaults to ``1.76``.

        Raises:
            ValueError: If ``color`` or ``width`` are less than zero.

        Returns:
            None:
        """
        super().__init__()

        self._location = location
        self._color = color
        self._transparent = transparent
        try:
            self._width = width.get_value_mm100()  # type: ignore
        except AttributeError:
            self._width = UnitConvert.convert_mm_mm100(width)  # type: ignore

        if self._color < 0:
            raise ValueError("color must be a positive number")

        if self._width < 0:
            raise ValueError("Width must be a positive number")

    # endregion init

    # region methods
    def __eq__(self, other: Any) -> bool:
        s2 = None
        if isinstance(other, ShadowStruct):
            s2 = other.get_uno_struct()
        elif getattr(other, "typeName", None) == "com.sun.star.table.ShadowFormat":
            s2 = other
        if s2:
            s1 = self.get_uno_struct()
            return (
                s1.Color == s2.Color
                and s1.IsTransparent == s2.IsTransparent
                and s1.Location == s2.Location
                and s1.ShadowWidth == s2.ShadowWidth
            )
        return False

    def get_uno_struct(self) -> ShadowFormat:
        """
        Gets UNO ``ShadowFormat`` from instance.

        Returns:
            ShadowFormat: ``ShadowFormat`` instance
        """
        return ShadowFormat(
            Location=self._location,  # type: ignore
            ShadowWidth=self._width,
            IsTransparent=self._transparent,
            Color=self._color,  # type: ignore
        )

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = (
                "com.sun.star.text.BaseFrame",
                "com.sun.star.text.TextEmbeddedObject",
                "com.sun.star.text.TextFrame",
                "com.sun.star.text.TextGraphicObject",
            )
        return self._supported_services_values

    def _on_modifying(self, source: Any, event: CancelEventArgs) -> None:
        if self._is_default_inst:
            raise ValueError("Modifying a default instance is not allowed")
        return super()._on_modifying(source, event)

    def _get_property_name(self) -> str:
        try:
            return self._property_name
        except AttributeError:
            self._property_name = "ShadowFormat"
        return self._property_name

    def get_attrs(self) -> Tuple[str, ...]:
        """
        Gets the attributes that are slated for change in the current instance

        Returns:
            Tuple(str, ...): Tuple of attributes
        """
        return (self._get_property_name(),)

    # region copy()
    @overload
    def copy(self: _TShadowStruct) -> _TShadowStruct: ...

    @overload
    def copy(self: _TShadowStruct, **kwargs) -> _TShadowStruct: ...

    def copy(self: _TShadowStruct, **kwargs) -> _TShadowStruct:
        """Gets a copy of instance as a new instance"""
        return self.__class__(
            location=self.prop_location,
            color=self.prop_color,
            transparent=self.prop_transparent,
            width=self.prop_width,
            **kwargs,
        )

    # endregion copy()

    # region apply()

    @overload
    def apply(self, obj: Any) -> None: ...

    @overload
    def apply(self, obj: Any, keys: Dict[str, str]) -> None: ...

    def apply(self, obj: Any, **kwargs) -> None:  # type: ignore
        """
        Applies style to object

        Args:
            obj (object): Object that contains a ``ShadowFormat`` property.
            keys: (Dict[str, str], optional): key map for properties.
                Can be ``prop`` which defaults to ``ShadowFormat``.

        Returns:
            None:
        """
        # sourcery skip: dict-assign-update-to-union
        keys = {"prop": self.property_name}
        if "keys" in kwargs:
            keys.update(kwargs["keys"])

        shadow = self.get_uno_struct()
        super().apply(obj=obj, override_dv={keys["prop"]: shadow})

    def _props_set(self, obj: Any, **kwargs: Any) -> None:
        try:
            super()._props_set(obj, **kwargs)
        except mEx.MultiError as e:
            mLo.Lo.print("Hatch.apply(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")

    # endregion apply()
    # endregion methods

    # region Static Methods

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TShadowStruct], obj: Any) -> _TShadowStruct: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TShadowStruct], obj: Any, **kwargs) -> _TShadowStruct: ...

    @classmethod
    def from_obj(cls: Type[_TShadowStruct], obj: Any, **kwargs) -> _TShadowStruct:
        """
        Gets instance from object

        Args:
            obj (object): UNO object

        Returns:
            Shadow: Instance from object
        """

        nu = cls(**kwargs)

        shadow = cast(ShadowFormat, mProps.Props.get(obj, nu.property_name))
        if shadow is None:
            return nu.empty.copy()

        return cls(
            location=shadow.Location,
            color=shadow.Color,  # type: ignore
            transparent=shadow.IsTransparent,
            width=UnitMM100(shadow.ShadowWidth).get_value_mm(),
            **kwargs,
        )

    # endregion from_obj()

    # region from_shadow()
    @overload
    @classmethod
    def from_uno_struct(cls: Type[_TShadowStruct], shadow: ShadowFormat) -> _TShadowStruct: ...

    @overload
    @classmethod
    def from_uno_struct(cls: Type[_TShadowStruct], shadow: ShadowFormat, **kwargs) -> _TShadowStruct: ...

    @classmethod
    def from_uno_struct(cls: Type[_TShadowStruct], shadow: ShadowFormat, **kwargs) -> _TShadowStruct:
        """
        Gets an instance

        Args:
            shadow (ShadowFormat): Shadow Format

        Returns:
            Shadow: Instance representing ``shadow``.
        """
        return cls(
            location=shadow.Location,
            color=shadow.Color,  # type: ignore
            transparent=shadow.IsTransparent,
            width=UnitMM100(shadow.ShadowWidth),
            **kwargs,
        )

    # endregion from_shadow()
    # endregion Static Methods

    # region style methods
    def fmt_location(self: _TShadowStruct, value: ShadowLocation) -> _TShadowStruct:
        """
        Gets a copy of instance with location set

        Args:
            value (ShadowLocation): Shadow location value

        Returns:
            Shadow: Shadow with location set
        """
        cp = self.copy()
        cp.prop_location = value
        return cp

    def fmt_color(self: _TShadowStruct, value: Color) -> _TShadowStruct:
        """
        Gets a copy of instance with color set

        Args:
            value (Color): color value

        Returns:
            Shadow: Shadow with color set
        """
        cp = self.copy()
        cp.prop_color = value
        return cp

    def fmt_transparent(self: _TShadowStruct, value: bool) -> _TShadowStruct:
        """
        Gets a copy of instance with transparency set

        Args:
            value (bool): transparency value

        Returns:
            Shadow: Shadow with transparency set
        """
        cp = self.copy()
        cp.prop_transparent = value
        return cp

    def fmt_width(self: _TShadowStruct, value: float | UnitT) -> _TShadowStruct:
        """
        Gets a copy of instance with width set

        Args:
            value (float): width value

        Returns:
            Shadow: Shadow with width set
        """
        cp = self.copy()
        cp.prop_width = value
        return cp

    # endregion style methods

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
    def prop_location(self) -> ShadowLocation:
        """Gets the location of the shadow."""
        return self._location

    @prop_location.setter
    def prop_location(self, value: ShadowLocation) -> None:
        self._location = value

    @property
    def prop_color(self) -> Color:
        """Gets the color value of the shadow."""
        return self._color

    @prop_color.setter
    def prop_color(self, value: Color) -> None:
        self._color = value

    @property
    def prop_transparent(self) -> bool:
        """Gets transparent value"""
        return self._transparent

    @prop_transparent.setter
    def prop_transparent(self, value: bool) -> None:
        self._transparent = value

    @property
    def prop_width(self) -> UnitMM:
        """Gets the size of the shadow (in mm units)"""
        return UnitMM.from_mm100(self._width)

    @prop_width.setter
    def prop_width(self, value: float | UnitT) -> None:
        try:
            self._width = value.get_value_mm100()  # type: ignore
        except AttributeError:
            self._width = UnitConvert.convert_mm_mm100(value)  # type: ignore

    @property
    def property_name(self) -> str:
        """
        Gets/Sets property name

        This is the name of the property that the side should be applied to. Such as ``LeftBorder``, ``RightBorder`` etc.
        """
        return self._get_property_name()

    @property_name.setter
    def property_name(self, value: str) -> None:
        self._set_property_name(value)

    @property
    def empty(self: _TShadowStruct) -> _TShadowStruct:  # type: ignore[misc]
        """Gets empty Shadow. When style is applied it remove any shadow."""
        try:
            return self._empty_inst
        except AttributeError:
            # pylint: disable=unexpected-keyword-arg
            # pylint: disable=protected-access
            self._empty_inst = self.__class__(
                location=ShadowLocation.NONE,
                transparent=False,
                color=8421504,
                width=1.76,
                _cattribs=self._get_internal_cattribs(),  # type: ignore
            )
            self._empty_inst._is_default_inst = True
        return self._empty_inst

    # endregion Properties
