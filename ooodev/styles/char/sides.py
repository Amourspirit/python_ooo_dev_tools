# region imports
from __future__ import annotations
from typing import cast, overload, TYPE_CHECKING

import uno
from ...exceptions import ex as mEx
from ...utils import props as mProps
from ...utils import info as mInfo
from ...utils import lo as mLo
from ..style_base import StyleBase
from ...events.args.key_val_cancel_args import KeyValCancelArgs
from ..structs import side
from ..structs.side import Side as Side, BorderLineStyleEnum as BorderLineStyleEnum

from ooo.dyn.table.border_line2 import BorderLine2

if TYPE_CHECKING:
    from com.sun.star.style import CharacterProperties  # service

# endregion imports


class Sides(StyleBase):
    """Character Border for use in styles."""

    _CHAR_BORDERS = ("CharTopBorder", "CharBottomBorder", "CharLeftBorder", "CharRightBorder")
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
            init_vals["CharTopBorder"] = border_side
            init_vals["CharBottomBorder"] = border_side
            init_vals["CharLeftBorder"] = border_side
            init_vals["CharRightBorder"] = border_side

        else:
            if not top is None:
                init_vals["CharTopBorder"] = top
            if not bottom is None:
                init_vals["CharBottomBorder"] = bottom
            if not left is None:
                init_vals["CharLeftBorder"] = left
            if not right is None:
                init_vals["CharRightBorder"] = right

        self._has_attribs = len(init_vals) > 0
        super().__init__(**init_vals)

    # endregion init

    # region methods

    # region apply_style()

    @overload
    def apply_style(self, obj: object) -> None:
        ...

    def apply_style(self, obj: object, **kwargs) -> None:
        """
        Applies Style to obj

        Args:
            obj (object): UNO object

        Returns:
            None:
        """
        if mInfo.Info.support_service(obj, "com.sun.star.style.CharacterProperties"):
            try:
                super().apply_style(obj)
            except mEx.MultiError as e:
                mLo.Lo.print(f"BorderChar.apply_style(): Unable to set Property")
                for err in e.errors:
                    mLo.Lo.print(f"  {err}")
        else:
            mLo.Lo.print('BorderChar.apply_style(): "com.sun.star.style.CharacterProperties" not supported')

    # endregion apply_style()

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

    @staticmethod
    def from_obj(obj: object) -> Sides:
        """
        Gets instance from object properties

        Args:
            obj (object): UNO object that has supports ``com.sun.star.style.CharacterProperties`` service.

        Raises:
            PropertyNotFoundError: If ``obj`` does not have ``TableBorder2`` property.

        Returns:
            BorderTable: Border Table.
        """
        if not mInfo.Info.support_service(obj, "com.sun.star.style.CharacterProperties"):
            mLo.Lo.print('BorderChar.apply_style(): "com.sun.star.style.CharacterProperties" not supported')
            return

        cp = cast("CharacterProperties", obj)
        empty = BorderLine2()
        bc = Sides()
        for attr in Sides._CHAR_BORDERS:
            b2 = cast(BorderLine2, getattr(cp, attr, empty))
            side = Side.from_border2(b2)
            bc._set(attr, side)
        return bc

    # endregion methods

    # region Properties

    @property
    def left(self) -> Side | None:
        """Gets left value"""
        return self._get("CharLeftBorder")

    @property
    def right(self) -> Side | None:
        """Gets right value"""
        return self._get("CharRightBorder")

    @property
    def top(self) -> Side | None:
        """Gets bottom value"""
        return self._get("CharTopBorder")

    @property
    def bottom(self) -> Side | None:
        """Gets bottom value"""
        return self._get("CharBottomBorder")

    # endregion Properties
