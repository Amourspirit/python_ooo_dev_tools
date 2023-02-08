"""
Module for Fill Properties Fill Color.

.. versionadded:: 0.9.0
"""
from __future__ import annotations
from typing import Tuple, overload

from ....events.args.cancel_event_args import CancelEventArgs
from ....exceptions import ex as mEx
from ....utils import props as mProps
from ....meta.static_prop import static_prop
from ....utils import lo as mLo
from ....utils.color import Color
from ...kind.format_kind import FormatKind
from ...style_base import StyleBase

from ooo.dyn.drawing.fill_style import FillStyle as FillStyle

# LibreOffice seems to have an unresolved bug with Background color.
# https://bugs.documentfoundation.org/show_bug.cgi?id=99125
# see Also: https://forum.openoffice.org/en/forum/viewtopic.php?p=417389&sid=17b21c173e4a420b667b45a2949b9cc5#p417389


class FillColor(StyleBase):
    """
    Class for Fill Properties Fill Color.

    .. versionadded:: 0.9.0
    """

    def __init__(self, color: Color = -1) -> None:
        """
        Constructor

        Args:
            color (Color, optional): FillColor Color

        Returns:
            None:
        """

        init_vals = {}
        if color >= 0:
            init_vals["FillColor"] = color
            init_vals["FillBackground"] = True
            init_vals["FillStyle"] = FillStyle.SOLID
        else:
            init_vals["FillColor"] = -1
            init_vals["FillBackground"] = False
            init_vals["FillStyle"] = FillStyle.NONE

        super().__init__(**init_vals)

    def _supported_services(self) -> Tuple[str, ...]:
        """
        Gets a tuple of supported services (``com.sun.star.drawing.FillProperties``,)

        Returns:
            Tuple[str, ...]: Supported services
        """
        return (
            "com.sun.star.drawing.FillProperties",
            "com.sun.star.beans.PropertySet",
            "com.sun.star.chart2.PageBackground",
        )

    def _on_modifing(self, event: CancelEventArgs) -> None:
        if self._is_default_inst:
            raise ValueError("Setting properties on a default instance is not allowed")
        return super()._on_setting(event)

    # region apply()

    @overload
    def apply(self, obj: object) -> None:
        ...

    def apply(self, obj: object, **kwargs) -> None:
        """
        Applies padding to ``obj``

        Args:
            obj (object): UNO object that supports ``com.sun.star.drawing.FillProperties`` service.

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

    @classmethod
    def from_obj(cls, obj: object) -> FillColor:
        """
        Gets instance from object

        Args:
            obj (object): UNO object that supports ``com.sun.star.drawing.FillProperties`` service.

        Raises:
            NotSupportedServiceError: If ``obj`` does not support  ``com.sun.star.drawing.FillProperties`` service.

        Returns:
            FillColor: ``FillColor`` instance that represents ``obj`` Fill Color properties.
        """

        nu = super(Color, cls).__new__(cls)
        nu.__init__()

        # inst = Color()
        if not nu._is_valid_obj(obj):
            raise mEx.NotSupportedError("Object does not suport conversion to color")

        color = mProps.Props.get(obj, "FillColor", None)

        inst = super(Color, cls).__new__(cls)

        if color is None:
            inst.__init__()
        else:
            inst.__init__(color=color)
        return inst

    def _is_valid_obj(self, obj: object) -> bool:
        valid = super()._is_valid_obj(obj)
        if valid:
            return True
        if mLo.Lo.is_uno_interfaces(obj, "com.sun.star.beans.XPropertySet"):
            return True
        return False

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        return FormatKind.TXT_CONTENT

    @property
    def prop_color(self) -> Color:
        """Gets/Sets color"""
        return self._get("FillColor")

    @prop_color.setter
    def prop_color(self, value: Color):
        if value >= 0:
            self._set("FillColor", value)
            self._set("FillBackground", True)
            self._set("FillStyle", FillStyle.SOLID)
        else:
            self._set("FillColor", -1)
            self._set("FillBackground", False)
            self._set("FillStyle", FillStyle.NONE)

    @static_prop
    def empty() -> FillColor:  # type: ignore[misc]
        """Gets FillColor empty. Static Property."""
        try:
            return FillColor._EMPTY_PROPS
        except AttributeError:
            FillColor._EMPTY_PROPS = FillColor()
            FillColor._EMPTY_PROPS._is_default_inst = True
        return FillColor._EMPTY_PROPS
