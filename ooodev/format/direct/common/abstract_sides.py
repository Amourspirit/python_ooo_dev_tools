"""
Module for managing character border side.

.. versionadded:: 0.9.0
"""
# region imports
from __future__ import annotations
from typing import Tuple, cast, overload, TYPE_CHECKING

import uno
from ....events.args.key_val_cancel_args import KeyValCancelArgs
from ....exceptions import ex as mEx
from ....utils import lo as mLo
from ...kind.format_kind import FormatKind
from ..structs import side
from ..structs.side import Side as Side, BorderLineStyleEnum as BorderLineStyleEnum
from ...style_base import StyleBase
from .border_props import BorderProps as BorderProps

if TYPE_CHECKING:
    try:
        from typing import Self
    except ImportError:
        from typing_extensions import Self

# endregion imports


class AbstractSides(StyleBase):
    """
    Character Border for use in styles.

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together properties.

    .. versionadded:: 0.9.0
    """

    # region init

    def __init__(
        self,
        left: Side | None = None,
        right: Side | None = None,
        top: Side | None = None,
        bottom: Side | None = None,
        border_side: Side | None = None,
    ) -> None:
        """
        Constructor

        Args:
            left (Side | None, optional): Determines the line style at the left edge.
            right (Side | None, optional): Determines the line style at the right edge.
            top (Side | None, optional): Determines the line style at the top edge.
            bottom (Side | None, optional): Determines the line style at the bottom edge.
            border_side (Side | None, optional): Determines the line style at the top, bottom, left, right edges. If this argument has a value then arguments ``top``, ``bottom``, ``left``, ``right`` are ignored
        """
        init_vals = {}
        if not border_side is None:
            init_vals[self._border.top] = border_side
            init_vals[self._border.bottom] = border_side
            init_vals[self._border.left] = border_side
            init_vals[self._border.right] = border_side

        else:
            if not top is None:
                init_vals[self._border.top] = top
            if not bottom is None:
                init_vals[self._border.bottom] = bottom
            if not left is None:
                init_vals[self._border.left] = left
            if not right is None:
                init_vals[self._border.right] = right

        self._has_attribs = len(init_vals) > 0
        super().__init__(**init_vals)

    # endregion init

    # region style methods
    # def style_left(self, value: Side | None = None) -> Sides:
    #     pass
    # endregion style methods
    # region methods

    def _supported_services(self) -> Tuple[str, ...]:
        """
        Gets a tuple of supported services (``com.sun.star.style.CharacterProperties``,)

        Returns:
            Tuple[str, ...]: Supported services
        """
        return ("com.sun.star.style.CharacterProperties",)

    # region apply()

    @overload
    def apply(self, obj: object) -> None:
        ...

    def apply(self, obj: object, **kwargs) -> None:
        """
        Applies Style to obj

        Args:
            obj (object): UNO object

        Returns:
            None:
        """
        try:
            super().apply(obj, **kwargs)
        except mEx.MultiError as e:
            mLo.Lo.print(f"{self.__class__.__name__}.apply(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")

    # endregion apply()

    def on_property_setting(self, event_args: KeyValCancelArgs):
        """
        Raise for each property that is set

        Args:
            event_args (KeyValueCancelArgs): Event Args
        """
        if event_args.has("sides_border2_set"):
            return
        side = cast(Side, event_args.value)
        event_args.value = side.get_border_line2()
        event_args.set("sides_border2_set", True)

    def __eq__(self, other: object) -> bool:
        if isinstance(other, AbstractSides):
            return (
                self.prop_bottom == other.prop_bottom
                and self.prop_left == other.prop_left
                and self.prop_right == other.prop_right
                and self.prop_top == other.prop_top
            )
        return False

    # endregion methods

    # region style methods
    def fmt_border_side(self, value: Side | None) -> AbstractSides:
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

    def fmt_top(self, value: Side | None) -> Self:
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

    def fmt_bottom(self, value: Side | None) -> Self:
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

    def fmt_left(self, value: Side | None) -> Self:
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

    def fmt_right(self, value: Side | None) -> Self:
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
        return FormatKind.CHAR

    @property
    def prop_left(self) -> Side | None:
        """Gets left value"""
        return self._get(self._border.left)

    @prop_left.setter
    def prop_left(self, value: Side | None) -> None:
        if value is None:
            self._remove(self._border.left)
            return
        self._set(self._border.left, value)

    @property
    def prop_right(self) -> Side | None:
        """Gets right value"""
        return self._get(self._border.right)

    @prop_right.setter
    def prop_right(self, value: Side | None) -> None:
        if value is None:
            self._remove(self._border.right)
            return
        self._set(self._border.right, value)

    @property
    def prop_top(self) -> Side | None:
        """Gets top value"""
        return self._get(self._border.top)

    @prop_top.setter
    def prop_top(self, value: Side | None) -> None:
        if value is None:
            self._remove(self._border.top)
            return
        self._set(self._border.top, value)

    @property
    def prop_bottom(self) -> Side | None:
        """Gets bottom value"""
        return self._get(self._border.bottom)

    @prop_bottom.setter
    def prop_bottom(self, value: Side | None) -> None:
        if value is None:
            self._remove(self._border.bottom)
            return
        self._set(self._border.bottom, value)

    @property
    def _border(self) -> BorderProps:
        try:
            return self.__border_properties
        except AttributeError:
            self.__border_properties = BorderProps(
                left="CharLeftBorder", top="CharTopBorder", right="CharRightBorder", bottom="CharBottomBorder"
            )
        return self.__border_properties

    # endregion Properties
