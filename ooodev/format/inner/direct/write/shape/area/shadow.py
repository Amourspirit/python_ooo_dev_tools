# region Imports
from __future__ import annotations
from typing import Any, Tuple, Type, TypeVar, cast, overload, TYPE_CHECKING
from enum import Enum

from ooodev.exceptions import ex as mEx
from ooodev.format.inner.common.props.shape_shadow_props import ShapeShadowProps
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleBase
from ooodev.loader import lo as mLo
from ooodev.units.unit_convert import UnitConvert
from ooodev.units.unit_mm import UnitMM
from ooodev.units.unit_pt import UnitPT
from ooodev.utils import props as mProps
from ooodev.utils.color import Color
from ooodev.utils.data_type.intensity import Intensity


if TYPE_CHECKING:
    from ooodev.units.unit_obj import UnitT

# endregion Imports

_TShadow = TypeVar("_TShadow", bound="Shadow")


class ShadowLocationKind(Enum):
    NONE = 0
    TOP_LEFT = 1
    TOP = 2
    TOP_RIGHT = 3
    RIGHT = 4
    BOTTOM_RIGHT = 5
    BOTTOM = 6
    BOTTOM_LEFT = 7
    LEFT = 8


class Shadow(StyleBase):
    """
    Shape Shadow

    .. seealso::

        - :ref:`help_writer_format_direct_shape_shadow`

    .. versionadded:: 0.9.0
    """

    # region Init
    def __init__(
        self,
        *,
        use_shadow: bool | None = None,
        location: ShadowLocationKind | None = None,
        color: Color | None = None,
        distance: float | UnitT | None = None,
        blur: int | UnitT | None = None,
        transparency: int | Intensity | None = None,
    ) -> None:
        """
        Constructor

        Args:
            use_shadow (bool, optional): Specifies if shadow is used.
            location (ShadowLocationKind , optional): Specifies the shadow location.
            color (Color , optional): Specifies shadow color.
            distance (float, UnitT , optional): Specifies shadow distance in ``mm`` units or :ref:`proto_unit_obj`.
            blur (int, UnitT, optional): Specifies shadow blur in ``pt`` units or :ref:`proto_unit_obj`.
            transparency (int , optional): Specifies shadow transparency value from ``0`` to ``100``.

        Returns:
            None:

        See Also:

            - :ref:`help_writer_format_direct_shape_shadow`
        """
        # shadow distance is stored in 1/100th mm.
        # shadow distance is stored in ShadowXDistance and ShadowYDistance depending on location value.
        super().__init__()
        self._location = location
        if use_shadow is not None:
            self.prop_use_shadow = use_shadow
        if color is not None:
            self.prop_color = color
        if distance is not None:
            self.prop_distance = distance
        if blur is not None:
            self.prop_blur = blur
        if transparency is not None:
            self.prop_transparency = transparency

    # endregion Init

    # region Overrides

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.drawing.ShadowProperties",)
        return self._supported_services_values

    def _props_set(self, obj: Any, **kwargs: Any) -> None:
        try:
            return super()._props_set(obj, **kwargs)
        except mEx.MultiError as e:
            mLo.Lo.print(f"{self.__class__.__name__}.apply(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")

    def copy(self: _TShadow) -> _TShadow:
        """Gets a copy of instance as a new instance"""
        cp = super().copy()
        cp._location = self._location
        return cp

    # endregion Overrides

    # region Internal Methods
    def _get_shadow_distance(self) -> int:
        """Returns distance in ``1/100th mm`` units."""
        if self._location is None:
            return 0

        def get_xy(x: bool = True) -> int:
            xy = self._props.dist_x if x else self._props.dist_y
            pv = cast(int, self._get(xy))
            return 0 if pv is None else abs(pv)

        if self._location == ShadowLocationKind.TOP_LEFT:
            # x or y, both negative values
            return get_xy()
        if self._location == ShadowLocationKind.TOP:
            # y, negative value
            return get_xy(False)
        if self._location == ShadowLocationKind.TOP_RIGHT:
            # x, y, x is positive, y is negative
            return get_xy()
        if self._location == ShadowLocationKind.RIGHT:
            # x, positive value
            return get_xy()
        if self._location == ShadowLocationKind.BOTTOM_RIGHT:
            # x or y, both are positive values
            return get_xy()
        if self._location == ShadowLocationKind.BOTTOM:
            # y, positive value
            return get_xy(False)
        if self._location == ShadowLocationKind.BOTTOM_LEFT:
            # x or y, x is negative, y is positive
            return get_xy()
        return get_xy() if self._location == ShadowLocationKind.LEFT else 0

    def _set_shadow_distance(self, value: int) -> None:
        """value is ``1/100th mm``"""

        val = abs(value)

        def set_xy(x: int, y: int) -> None:
            self._set(self._props.dist_x, x)
            self._set(self._props.dist_y, y)

        if self._location is None:
            set_xy(0, 0)
            return

        if self._location == ShadowLocationKind.TOP_LEFT:
            set_xy(-val, -val)
        elif self._location == ShadowLocationKind.TOP:
            set_xy(0, -val)
        elif self._location == ShadowLocationKind.TOP_RIGHT:
            set_xy(val, -val)
        elif self._location == ShadowLocationKind.RIGHT:
            set_xy(val, 0)
        elif self._location == ShadowLocationKind.BOTTOM_RIGHT:
            set_xy(val, val)
        elif self._location == ShadowLocationKind.BOTTOM:
            set_xy(0, val)
        elif self._location == ShadowLocationKind.BOTTOM_LEFT:
            set_xy(-val, val)
        elif self._location == ShadowLocationKind.LEFT:
            set_xy(-val, 0)
        else:
            set_xy(0, 0)

    def _get_location_from_props(self) -> ShadowLocationKind | None:
        x = cast(int, self._get(self._props.dist_x))
        y = cast(int, self._get(self._props.dist_y))
        if x is None or y is None:
            return None
        if x < 0 and y < 0:
            return ShadowLocationKind.TOP_LEFT
        if x == 0 and y < 0:
            return ShadowLocationKind.TOP
        if x > 0:
            if y < 0:
                return ShadowLocationKind.TOP_RIGHT
            if y == 0:
                return ShadowLocationKind.RIGHT
            if y > 0:
                return ShadowLocationKind.BOTTOM_RIGHT
        if x == 0 and y > 0:
            return ShadowLocationKind.BOTTOM
        if x < 0:
            if y > 0:
                return ShadowLocationKind.BOTTOM_LEFT
            if y == 0:
                return ShadowLocationKind.LEFT
        return ShadowLocationKind.NONE

    # endregion Internal Methods

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TShadow], obj: Any) -> _TShadow: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TShadow], obj: Any, **kwargs) -> _TShadow: ...

    @classmethod
    def from_obj(cls: Type[_TShadow], obj: Any, **kwargs) -> _TShadow:
        """
        Gets instance from object

        Args:
            obj (object): UNO Object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            Shadow: Instance that represents Shadow settings.
        """
        # pylint: disable=protected-access
        # this nu is only used to get Property Name

        inst = cls(**kwargs)
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')
        for prop in inst._props:
            inst._set(prop, mProps.Props.get(obj, prop))
        inst._location = inst._get_location_from_props()
        inst.set_update_obj(obj)
        return inst

    # endregion from_obj()

    # region style methods
    def fmt_use_shadow(self: _TShadow, value: bool | None = None) -> _TShadow:
        """
        Get copy of instance with use shadow set or removed.

        Args:
            value (bool, optional): Specifies if shadow is used.

        Returns:
            Shadow: Shadow with style added or removed
        """
        cp = self.copy()
        cp.prop_use_shadow = value
        return cp

    def fmt_location(self: _TShadow, value: ShadowLocationKind | None = None) -> _TShadow:
        """
        Get copy of instance with use shadow location or removed.

        Args:
            value (ShadowLocationKind, optional): Specifies the shadow location.

        Returns:
            Shadow: Shadow with style added or removed
        """
        cp = self.copy()
        cp.prop_location = value
        return cp

    def fmt_color(self: _TShadow, value: Color | None = None) -> _TShadow:
        """
        Get copy of instance with use shadow color or removed.

        Args:
            value (:py:data:`~.utils.color.Color`, optional): Specifies shadow color.

        Returns:
            Shadow: Shadow with style added or removed
        """
        cp = self.copy()
        cp.prop_color = value
        return cp

    def fmt_distance(self: _TShadow, value: float | UnitT | None = None) -> _TShadow:
        """
        Get copy of instance with use shadow distance or removed.

        Args:
            value (float, UnitT, optional): Specifies shadow distance in ``mm`` units or :ref:`proto_unit_obj`.

        Returns:
            Shadow: Shadow with style added or removed
        """
        cp = self.copy()
        cp.prop_distance = value
        return cp

    def fmt_blur(self: _TShadow, value: int | UnitT | None = None) -> _TShadow:
        """
        Get copy of instance with use shadow blur or removed.

        Args:
            value (int, UnitT, optional): Specifies shadow blur in ``pt`` units or :ref:`proto_unit_obj`.

        Returns:
            Shadow: Shadow with style added or removed
        """
        cp = self.copy()
        cp.prop_blur = value
        return cp

    def fmt_transparency(self: _TShadow, value: int | Intensity | None = None) -> _TShadow:
        """
        Get copy of instance with use shadow transparency or removed.

        Args:
            value (int, Intensity, optional): Specifies shadow transparency value from ``0`` to ``100``.

        Returns:
            Shadow: Shadow with style added or removed
        """
        cp = self.copy()
        cp.prop_transparency = value
        return cp

    # endregion style methods

    # region style properties
    @property
    def use_shadow(self: _TShadow) -> _TShadow:
        """Gets instance with use shadow set to ``True``."""
        cp = self.copy()
        cp.prop_use_shadow = True
        return cp

    @property
    def top_left(self: _TShadow) -> _TShadow:
        """Gets instance with shadow location set to top left."""
        cp = self.copy()
        cp.prop_location = ShadowLocationKind.TOP_LEFT
        return cp

    @property
    def top(self: _TShadow) -> _TShadow:
        """Gets instance with shadow location set to top."""
        cp = self.copy()
        cp.prop_location = ShadowLocationKind.TOP
        return cp

    @property
    def top_right(self: _TShadow) -> _TShadow:
        """Gets instance with shadow location set to top right."""
        cp = self.copy()
        cp.prop_location = ShadowLocationKind.TOP_RIGHT
        return cp

    @property
    def right(self: _TShadow) -> _TShadow:
        """Gets instance with shadow location set to right."""
        cp = self.copy()
        cp.prop_location = ShadowLocationKind.RIGHT
        return cp

    @property
    def bottom_right(self: _TShadow) -> _TShadow:
        """Gets instance with shadow location set to bottom right."""
        cp = self.copy()
        cp.prop_location = ShadowLocationKind.BOTTOM_RIGHT
        return cp

    @property
    def bottom(self: _TShadow) -> _TShadow:
        """Gets instance with shadow location set to bottom."""
        cp = self.copy()
        cp.prop_location = ShadowLocationKind.BOTTOM
        return cp

    @property
    def bottom_left(self: _TShadow) -> _TShadow:
        """Gets instance with shadow location set to bottom left."""
        cp = self.copy()
        cp.prop_location = ShadowLocationKind.BOTTOM_LEFT
        return cp

    @property
    def left(self: _TShadow) -> _TShadow:
        """Gets instance with shadow location set to left."""
        cp = self.copy()
        cp.prop_location = ShadowLocationKind.LEFT
        return cp

    # endregion style properties

    # region Prop Properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.SHAPE
        return self._format_kind_prop

    @property
    def prop_use_shadow(self) -> bool | None:
        """Gets/Sets Shadow using value"""
        return self._get(self._props.use)

    @prop_use_shadow.setter
    def prop_use_shadow(self, value: bool | None) -> None:
        if value is None:
            self._remove(self._props.use)
            return
        self._set(self._props.use, value)

    @property
    def prop_location(self) -> ShadowLocationKind | None:
        """Gets/Sets Shadow location value"""
        return self._location

    @prop_location.setter
    def prop_location(self, value: ShadowLocationKind | None) -> None:
        self._location = value
        self._set_shadow_distance(self._get_shadow_distance())

    @property
    def prop_color(self) -> Color | None:
        """Gets/Sets Shadow color value"""
        return self._get(self._props.color)

    @prop_color.setter
    def prop_color(self, value: Color | None) -> None:
        if value is None:
            self._remove(self._props.color)
            return
        self._set(self._props.color, value)

    @property
    def prop_distance(self) -> UnitMM | None:
        """Gets/Sets Shadow distance value. Return value is in ``mm`` units."""
        return UnitMM.from_mm100(self._get_shadow_distance())

    @prop_distance.setter
    def prop_distance(self, value: float | UnitT | None) -> None:
        if value is None:
            self._remove(self._props.dist_x)
            self._remove(self._props.dist_y)
            return
        try:
            self._set_shadow_distance(value.get_value_mm100())  # type: ignore
        except AttributeError:
            self._set_shadow_distance(UnitConvert.convert_mm_mm100(value))  # type: ignore

    @property
    def prop_blur(self) -> UnitPT | None:
        """Gets/Sets Shadow blur value. Return value is in ``pt`` (points) units."""
        pv = cast(int, self._get(self._props.blur))
        return None if pv is None else UnitPT(round(UnitConvert.convert_mm100_pt(pv)))

    @prop_blur.setter
    def prop_blur(self, value: int | UnitT | None) -> None:
        if value is None:
            self._remove(self._props.blur)
            return
        try:
            self._set(self._props.blur, value.get_value_mm100())  # type: ignore
        except AttributeError:
            self._set(self._props.blur, UnitConvert.convert_pt_mm100(value))  # type: ignore

    @property
    def prop_transparency(self) -> Intensity | None:
        """Gets/Sets Shadow transparency value"""
        pv = cast(int, self._get(self._props.transparence))
        return None if pv is None else Intensity(pv)

    @prop_transparency.setter
    def prop_transparency(self, value: int | Intensity | None) -> None:
        if value is None:
            self._remove(self._props.transparence)
            return

        self._set(self._props.transparence, Intensity(int(value)).value)

    @property
    def _props(self) -> ShapeShadowProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = ShapeShadowProps(
                use="Shadow",
                blur="ShadowBlur",
                color="ShadowColor",
                transparence="ShadowTransparence",
                dist_x="ShadowXDistance",
                dist_y="ShadowYDistance",
            )
        return self._props_internal_attributes

    # endregion Prop Properties
