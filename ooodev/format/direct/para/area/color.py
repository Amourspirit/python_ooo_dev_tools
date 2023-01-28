"""
Module for Paragraph Fill Color.

.. versionadded:: 0.9.0
"""
from __future__ import annotations
from typing import Tuple, overload

from .....meta.static_prop import static_prop
from .....utils import lo as mLo
from .....utils import props as mProps
from .....exceptions import ex as mEx
from .....utils import color as mColor
from ....kind.format_kind import FormatKind
from ....style_base import StyleBase
from ooo.dyn.drawing.fill_style import FillStyle

# LibreOffice seems to have an unresolved bug with Background color.
# https://bugs.documentfoundation.org/show_bug.cgi?id=99125
# see Also: https://forum.openoffice.org/en/forum/viewtopic.php?p=417389&sid=17b21c173e4a420b667b45a2949b9cc5#p417389
# The solution to these issues is to apply FillColor to Paragraph cursors TextParagraph.


class Color(StyleBase):
    """
    Paragraph Fill Coloring

    .. versionadded:: 0.9.0
    """

    _DEFAULT = None

    def __init__(self, color: mColor.Color = -1) -> None:
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
            init_vals["FillStyle"] = FillStyle.SOLID
        else:
            init_vals["FillColor"] = -1
            init_vals["FillStyle"] = FillStyle.NONE

        super().__init__(**init_vals)

    def _supported_services(self) -> Tuple[str, ...]:
        return (
            "com.sun.star.drawing.FillProperties",
            "com.sun.star.text.TextContent",
        )

    # region apply()

    @overload
    def apply(self, obj: object) -> None:
        ...

    def apply(self, obj: object, **kwargs) -> None:
        """
        Applies padding to ``obj``

        Args:
            obj (object): UNO object that supports ``com.sun.star.style.ParagraphProperties`` service.

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

    def dispatch_reset(self) -> None:
        """
        Resets the cursor at is current position/selection to remove any Fill Color Formatting.

        Returns:
            None:
        """
        mLo.Lo.dispatch_cmd("BackgroundColor", mProps.Props.make_props(BackgroundColor=-1))
        mLo.Lo.dispatch_cmd("Escape")

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        return FormatKind.PARA | FormatKind.FILL

    @property
    def prop_color(self) -> mColor.Color:
        """Gets/Sets color"""
        return self._get("FillColor")

    @prop_color.setter
    def prop_color(self, value: mColor.Color):
        if self is Color.default:
            raise ValueError("Properties of FillColor.default can not be changed.")
        if value >= 0:
            self._set("FillColor", value)
            self._set("FillStyle", FillStyle.SOLID)
        else:
            self._set("FillColor", -1)
            self._set("FillStyle", FillStyle.NONE)

    @static_prop
    def default() -> Color:  # type: ignore[misc]
        """Gets FillColor empty. Static Property."""
        if Color._DEFAULT is None:
            Color._DEFAULT = Color(-1)
        return Color._DEFAULT