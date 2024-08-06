"""
Module for managing character border side.

.. versionadded:: 0.9.0
"""

# region imports
from __future__ import annotations
from typing import Any, Tuple, cast, overload, Type, TypeVar
import uno
from ooo.dyn.table.border_line2 import BorderLine2
from ooodev.exceptions import ex as mEx
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.direct.structs.side import Side as Side
from ooodev.format.inner.style_base import StyleMulti
from ooodev.format.inner.common.props.border_props import BorderProps as BorderProps


# endregion imports

_TAbstractSides = TypeVar("_TAbstractSides", bound="AbstractSides")


class AbstractSides(StyleMulti):
    """
    Character Border for use in styles.

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together properties.

    .. versionadded:: 0.9.0
    """

    # region init

    def __init__(
        self,
        *,
        left: Side | None = None,
        right: Side | None = None,
        top: Side | None = None,
        bottom: Side | None = None,
        all: Side | None = None,
    ) -> None:
        """
        Constructor

        Args:
            left (Side, optional): Determines the line style at the left edge.
            right (Side, optional): Determines the line style at the right edge.
            top (Side, optional): Determines the line style at the top edge.
            bottom (Side, optional): Determines the line style at the bottom edge.
            all (Side, optional): Determines the line style at the top, bottom, left, right edges.
                If this argument has a value then arguments ``top``, ``bottom``, ``left``, ``right`` are ignored
        """
        super().__init__()
        if all is not None:
            self.prop_left = all
            self.prop_right = all
            self.prop_top = all
            self.prop_bottom = all
        else:
            if top is not None:
                self.prop_top = top
            if bottom is not None:
                self.prop_bottom = bottom
            if left is not None:
                self.prop_left = left
            if right is not None:
                self.prop_right = right

    # endregion init

    # region style methods
    # def style_left(self, value: Side | None = None) -> Sides:
    #     pass
    # endregion style methods

    # region internal methods
    def _get_property_side(self, side: Side, prop: str) -> Side:
        return side.copy(
            _cattribs={
                "_property_name": prop,
                "_supported_services_values": self._supported_services(),
            }
        )

    # endregion internal methods

    # region Overloads

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.style.CharacterProperties",)
        return self._supported_services_values

    # endregion Overloads

    # region Dunder Methods
    def __eq__(self, other: object) -> bool:
        if isinstance(other, AbstractSides):
            return (
                self.prop_bottom == other.prop_bottom
                and self.prop_left == other.prop_left
                and self.prop_right == other.prop_right
                and self.prop_top == other.prop_top
            )
        return False

    # endregion Dunder Methods

    # region Static Methods

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TAbstractSides], obj: Any) -> _TAbstractSides: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TAbstractSides], obj: Any, **kwargs) -> _TAbstractSides: ...

    @classmethod
    def from_obj(cls: Type[_TAbstractSides], obj: Any, **kwargs) -> _TAbstractSides:
        """
        Gets instance from object properties

        Args:
            obj (object): UNO object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            Sides: Instance that represents ``BorderLine2``.
        """
        # pylint: disable=protected-access
        inst = cls(**kwargs)
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')

        empty = BorderLine2()
        b2 = cast(BorderLine2, getattr(obj, inst._props.bottom, empty))
        inst.prop_bottom = Side.from_uno_struct(b2)

        b2 = cast(BorderLine2, getattr(obj, inst._props.left, empty))
        inst.prop_left = Side.from_uno_struct(b2)

        b2 = cast(BorderLine2, getattr(obj, inst._props.top, empty))
        inst.prop_top = Side.from_uno_struct(b2)

        b2 = cast(BorderLine2, getattr(obj, inst._props.right, empty))
        inst.prop_right = Side.from_uno_struct(b2)
        inst.set_update_obj(obj)
        return inst

    # endregion from_obj()
    # endregion Static Methods

    # region style methods
    def fmt_border_side(self: _TAbstractSides, value: Side | None) -> _TAbstractSides:
        """
        Gets copy of instance with left, right, top, bottom sides set or removed

        Args:
            value (Side | None): Side value

        Returns:
            Sides: Sides instance
        """
        cp = self.copy()
        cp.prop_top = value
        cp.prop_bottom = value
        cp.prop_left = value
        cp.prop_right = value
        return cp

    def fmt_top(self: _TAbstractSides, value: Side | None) -> _TAbstractSides:
        """
        Gets a copy of instance with top side set or removed

        Args:
            value (Side | None): Side value

        Returns:
            Sides: Sides instance
        """
        cp = self.copy()
        cp.prop_top = value
        return cp

    def fmt_bottom(self: _TAbstractSides, value: Side | None) -> _TAbstractSides:
        """
        Gets a copy of instance with bottom side set or removed

        Args:
            value (Side | None): Side value

        Returns:
            Sides: Sides instance
        """
        cp = self.copy()
        cp.prop_bottom = value
        return cp

    def fmt_left(self: _TAbstractSides, value: Side | None) -> _TAbstractSides:
        """
        Gets a copy of instance with left side set or removed

        Args:
            value (Side | None): Side value

        Returns:
            Sides: Sides instance
        """
        cp = self.copy()
        cp.prop_left = value
        return cp

    def fmt_right(self: _TAbstractSides, value: Side | None) -> _TAbstractSides:
        """
        Gets a copy of instance with right side set or removed

        Args:
            value (Side | None): Side value

        Returns:
            Sides: Sides instance
        """
        cp = self.copy()
        cp.prop_right = value
        return cp

    # endregion style methods

    # region Properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.CHAR
        return self._format_kind_prop

    @property
    def prop_left(self) -> Side | None:
        """Gets left value"""
        try:
            return self._prop_left  # type: ignore
        except AttributeError:
            val = self._get_style("left")
            self._prop_left = None if val is None else val.style
        return self._prop_left  # type: ignore

    @prop_left.setter
    def prop_left(self, value: Side | None) -> None:
        self._del_attribs("_prop_left")
        if value is None:
            self._remove_style("left")
            return
        inst = self._get_property_side(value, self._props.left)
        self._set_style("left", inst, *inst.get_attrs())

    @property
    def prop_right(self) -> Side | None:
        """Gets right value"""
        try:
            return self._prop_right  # type: ignore
        except AttributeError:
            val = self._get_style("right")
            self._prop_right = None if val is None else val.style
        return self._prop_right  # type: ignore

    @prop_right.setter
    def prop_right(self, value: Side | None) -> None:
        self._del_attribs("_prop_right")
        if value is None:
            self._remove_style("right")
            return
        inst = self._get_property_side(value, self._props.right)
        self._set_style("right", inst, *inst.get_attrs())

    @property
    def prop_top(self) -> Side | None:
        """Gets top value"""
        try:
            return self._prop_top  # type: ignore
        except AttributeError:
            val = self._get_style("top")
            self._prop_top = None if val is None else val.style
        return self._prop_top  # type: ignore

    @prop_top.setter
    def prop_top(self, value: Side | None) -> None:
        self._del_attribs("_prop_top")
        if value is None:
            self._remove_style("top")
            return
        inst = self._get_property_side(value, self._props.top)
        self._set_style("top", inst, *inst.get_attrs())

    @property
    def prop_bottom(self) -> Side | None:
        """Gets bottom value"""
        try:
            return self._prop_bottom  # type: ignore
        except AttributeError:
            val = self._get_style("bottom")
            self._prop_bottom = None if val is None else val.style
        return self._prop_bottom  # type: ignore

    @prop_bottom.setter
    def prop_bottom(self, value: Side | None) -> None:
        self._del_attribs("_prop_bottom")
        if value is None:
            self._remove_style("bottom")
            return
        inst = self._get_property_side(value, self._props.bottom)
        self._set_style("bottom", inst, *inst.get_attrs())

    @property
    def _props(self) -> BorderProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = BorderProps(
                left="CharLeftBorder", top="CharTopBorder", right="CharRightBorder", bottom="CharBottomBorder"
            )
        return self._props_internal_attributes

    # endregion Properties
