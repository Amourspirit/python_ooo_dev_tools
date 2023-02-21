"""
Module for Shadow format (``ShadowFormat``) struct.

.. versionadded:: 0.9.0
"""
# region imports
from __future__ import annotations
from typing import Any, Dict, Tuple, Type, cast, overload, TypeVar

import uno
from ....events.event_singleton import _Events
from ....meta.static_prop import static_prop
from ....utils import props as mProps
from ....utils import lo as mLo
from ....utils.color import Color, StandardColor
from ....exceptions import ex as mEx
from ...kind.format_kind import FormatKind
from ...style_base import StyleBase, CancelEventArgs
from ....utils.unit_convert import UnitConvert, Length

from ooo.dyn.table.shadow_format import ShadowFormat as ShadowFormat
from ooo.dyn.table.shadow_location import ShadowLocation as ShadowLocation

_TShadowStruct = TypeVar(name="_TShadowStruct", bound="ShadowStruct")

# endregion imports
class ShadowStruct(StyleBase):
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
        width: float = 1.76,
    ) -> None:
        """
        Constructor

        Args:
            location (ShadowLocation, optional): contains the location of the shadow. Default to ``ShadowLocation.BOTTOM_RIGHT``.
            color (Color, optional):contains the color value of the shadow. Defaults to ``StandardColor.GRAY``.
            transparent (bool, optional): Shadow transparency. Defaults to False.
            width (float, optional): contains the size of the shadow (in mm units). Defaults to ``1.76``.

        Raises:
            ValueError: If ``color`` or ``width`` are less than zero.
        """
        if color < 0:
            raise ValueError("color must be a positive number")
        if width < 0:
            raise ValueError("Width must be a postivie number")

        self._location = location
        self._color = color
        self._transparent = transparent
        self._width: int = round(UnitConvert.convert(num=width, frm=Length.MM, to=Length.MM100))

        super().__init__()

    # endregion init

    # region methods
    def __eq__(self, other: object) -> bool:
        s2: ShadowFormat = None
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
            Location=self._location,
            ShadowWidth=self._width,
            IsTransparent=self._transparent,
            Color=self._color,
        )

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ()
        return self._supported_services_values

    def _on_modifing(self, event: CancelEventArgs) -> None:
        if self._is_default_inst:
            raise ValueError("Modifying a default instance is not allowed")
        return super()._on_modifing(event)

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
            Tuple(str, ...): Tuple of attribures
        """
        return (self._get_property_name(),)

    def copy(self: _TShadowStruct) -> _TShadowStruct:
        nu = self.__class__(
            location=self.prop_width, color=self.prop_color, transparent=self.prop_transparent, width=self.prop_width
        )
        if self._dv:
            nu._update(self._dv)
        return nu

    # region apply()

    @overload
    def apply(self, obj: object) -> None:
        ...

    @overload
    def apply(self, obj: object, keys: Dict[str, str]) -> None:
        ...

    def apply(self, obj: object, **kwargs) -> None:
        """
        Applies style to object

        Args:
            obj (object): Object that contains a ``ShadowFormat`` property.
            keys: (Dict[str, str], optional): key map for properties.
                Can be ``prop`` which defaults to ``ShadowFormat``.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.format_named_event.FormatNamedEvent.STYLE_APPLYING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.format_named_event.FormatNamedEvent.STYLE_APPLYED` :eventref:`src-docs-event`

        Returns:
            None:
        """
        keys = {"prop": self._get_property_name()}
        if "keys" in kwargs:
            keys.update(kwargs["keys"])

        shadow = self.get_uno_struct()
        super().apply(obj=obj, override_dv={keys["prop"]: shadow})

    def _props_set(self, obj: object, **kwargs: Any) -> None:
        try:
            super()._props_set(obj, **kwargs)
        except mEx.MultiError as e:
            mLo.Lo.print(f"Hatch.apply(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")

    # endregion apply()

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TShadowStruct], obj: object) -> _TShadowStruct:
        ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TShadowStruct], obj: object, **kwargs) -> _TShadowStruct:
        ...

    @classmethod
    def from_obj(cls: Type[_TShadowStruct], obj: object, **kwargs) -> _TShadowStruct:
        """
        Gets instance from object

        Args:
            obj (object): UNO object

        Returns:
            Shadow: Instance from object
        """

        # this nu is only used to get Property Name
        nu = cls(**kwargs)

        shadow = cast(ShadowFormat, mProps.Props.get(obj, nu._get_property_name()))
        if shadow is None:
            return cls.empty.copy()
        width = UnitConvert.convert(num=shadow.ShadowWidth, frm=Length.MM100, to=Length.MM)

        return cls(
            location=shadow.Location, color=shadow.Color, transparent=shadow.IsTransparent, width=width, **kwargs
        )

    # endregion from_obj()

    # region from_shadow()
    @overload
    @classmethod
    def from_shadow(cls: Type[_TShadowStruct], shadow: ShadowFormat) -> _TShadowStruct:
        ...

    @overload
    @classmethod
    def from_shadow(cls: Type[_TShadowStruct], shadow: ShadowFormat, **kwargs) -> _TShadowStruct:
        ...

    @classmethod
    def from_shadow(cls: Type[_TShadowStruct], shadow: ShadowFormat, **kwargs) -> _TShadowStruct:
        """
        Gets an instance

        Args:
            shadow (ShadowFormat): Shadow Format

        Returns:
            Shadow: Instance representing ``shadow``.
        """
        width = UnitConvert.convert(num=shadow.ShadowWidth, frm=Length.MM100, to=Length.MM)
        return cls(
            location=shadow.Location, color=shadow.Color, transparent=shadow.IsTransparent, width=width, **kwargs
        )

    # endregion from_shadow()
    # endregion methods

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

    def fmt_width(self: _TShadowStruct, value: float) -> _TShadowStruct:
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
    def prop_width(self) -> float:
        """Gets the size of the shadow (in mm units)"""
        return UnitConvert.convert(num=self._width, frm=Length.MM100, to=Length.MM)

    @prop_width.setter
    def prop_width(self, value: float) -> None:
        self._width = round(UnitConvert.convert(num=value, frm=Length.MM, to=Length.MM100))

    @static_prop
    def empty() -> ShadowStruct:  # type: ignore[misc]
        """Gets empty Shadow. Static Property. when style is applied it remove any shadow."""
        try:
            return ShadowStruct._EMPTY_INST
        except AttributeError:
            ShadowStruct._EMPTY_INST = ShadowStruct(
                location=ShadowLocation.NONE, transparent=False, color=8421504, width=1.76
            )
            ShadowStruct._EMPTY_INST._is_default_inst = True
        return ShadowStruct._EMPTY_INST

    # endregion Properties
