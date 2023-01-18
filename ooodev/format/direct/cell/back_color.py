"""
Module for Cell Properties Cell Back Color.

.. versionadded:: 0.9.0
"""
from __future__ import annotations
from typing import Tuple, overload

from ....exceptions import ex as mEx
from ....utils import props as mProps
from ....meta.static_prop import static_prop
from ....utils import lo as mLo
from ....utils.color import Color
from ...kind.format_kind import FormatKind
from ...style_base import StyleBase

# LibreOffice seems to have an unresolved bug with Background color.
# https://bugs.documentfoundation.org/show_bug.cgi?id=99125
# see Also: https://forum.openoffice.org/en/forum/viewtopic.php?p=417389&sid=17b21c173e4a420b667b45a2949b9cc5#p417389


class BackColor(StyleBase):
    """
    Class for Cell Properties Back Color.

    .. versionadded:: 0.9.0
    """

    _EMPTY = None

    def __init__(self, color: Color = -1) -> None:
        """
        Constructor

        Args:
            color (Color, optional): Color

        Returns:
            None:
        """

        init_vals = {}
        if color >= 0:
            init_vals["CellBackColor"] = color
            init_vals["IsCellBackgroundTransparent"] = False
        else:
            init_vals["CellBackColor"] = -1
            init_vals["IsCellBackgroundTransparent"] = True

        super().__init__(**init_vals)

    def _supported_services(self) -> Tuple[str, ...]:
        """
        Gets a tuple of supported services (``com.sun.star.table.CellProperties``,)

        Returns:
            Tuple[str, ...]: Supported services
        """
        return ("com.sun.star.table.CellProperties",)

    # region apply()

    @overload
    def apply(self, obj: object) -> None:
        ...

    def apply(self, obj: object, **kwargs) -> None:
        """
        Applies padding to ``obj``

        Args:
            obj (object): UNO object that supports ``com.sun.star.table.CellProperties`` service.

        Returns:
            None:
        """
        try:
            super().apply(obj, **kwargs)
        except mEx.MultiError as e:
            mLo.Lo.print(f"{self.__class__}.apply_style(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")

    # endregion apply()

    @staticmethod
    def from_obj(obj: object) -> BackColor:
        """
        Gets instance from object

        Args:
            obj (object): UNO object that supports ``com.sun.star.table.CellProperties`` service.

        Raises:
            NotSupportedServiceError: If ``obj`` does not support  ``com.sun.star.table.CellProperties`` service.

        Returns:
            BackColor: ``BackColor`` instance that represents ``obj`` Back Color properties.
        """
        inst = BackColor()
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedServiceError(inst._supported_services()[0])

        color = mProps.Props.get(obj, "CellBackColor", None)
        bg = mProps.Props.get(obj, "IsCellBackgroundTransparent", None)
        if not color is None:
            inst._set("CellBackColor", int(color))
        if not bg is None:
            inst._set("IsCellBackgroundTransparent", bool(bg))

        return inst

    # region set styles

    # endregion set styles

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        return FormatKind.CELL

    @property
    def prop_color(self) -> Color:
        """Gets/Sets color"""
        return self._get("CellBackColor")

    @prop_color.setter
    def prop_color(self, value: Color):
        if value >= 0:
            self._set("CellBackColor", value)
            self._set("IsCellBackgroundTransparent", False)
        else:
            self._set("CellBackColor", -1)
            self._set("IsCellBackgroundTransparent", True)

    @property
    def prop_is_bg_transparent(self) -> bool:
        """Gets/Sets if Background color is Transparent"""
        return self._get("IsCellBackgroundTransparent")

    @prop_is_bg_transparent.setter
    def prop_is_bg_transparent(self, value: bool):
        self._set("IsCellBackgroundTransparent", value)

    @static_prop
    def empty() -> BackColor:  # type: ignore[misc]
        """Gets BackColor empty. Static Property."""
        if BackColor._EMPTY is None:
            BackColor._EMPTY = BackColor()
        return BackColor._EMPTY
