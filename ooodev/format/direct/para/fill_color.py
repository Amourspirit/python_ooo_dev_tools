"""
Module for Paragraph Fill Color.

.. versionadded:: 0.9.0
"""
from __future__ import annotations
from typing import Tuple, overload

from ....exceptions import ex as mEx
from ....meta.static_prop import static_prop
from ....utils import lo as mLo
from ....utils import props as mProps
from ....events.event_singleton import _Events
from ....utils.color import Color
from ...kind.format_kind import FormatKind
from ...style_base import StyleBase, CancelEventArgs, EventArgs, FormatNamedEvent

# LibreOffice seems to have an unresolved bug with Background color.
# https://bugs.documentfoundation.org/show_bug.cgi?id=99125
# see Also: https://forum.openoffice.org/en/forum/viewtopic.php?p=417389&sid=17b21c173e4a420b667b45a2949b9cc5#p417389


class FillColor(StyleBase):
    """
    Paragraph Fill Coloring

    Warning:
        This class uses dispatch commands and is not suitable for use in headless mode.

    .. versionadded:: 0.9.0
    """

    _EMPTY = None

    def __init__(self, color: Color = -1) -> None:
        """
        Constructor

        Args:
            color (Color, optional): FillColor Color

        Returns:
            None:
        """
        if mLo.Lo.bridge_connector.headless:
            mLo.Lo.print("Warning! FillColor class is not suitable in Headless mode.")

        init_vals = {}
        if color >= 0:
            init_vals["FillColor"] = color
            init_vals["FillBackground"] = False
        else:
            init_vals["FillColor"] = -1
            init_vals["FillBackground"] = True

        super().__init__(**init_vals)

    def _supported_services(self) -> Tuple[str, ...]:
        """
        Gets a tuple of supported services (``com.sun.star.style.ParagraphProperties``,)

        Returns:
            Tuple[str, ...]: Supported services
        """
        return ("com.sun.star.style.ParagraphProperties",)

    # region apply()

    @overload
    def apply(self, obj: object) -> None:
        ...

    def apply(self, obj: object, **kwargs) -> None:
        """
        Applies padding to ``obj``

        Args:
            obj (object): UNO object that supports ``com.sun.star.style.ParagraphProperties`` service.
            kwargs (Any, optional): Expandable list of key value pairs that may be used in child classes.

        Returns:
            None:
        """
        cargs = CancelEventArgs(source=f"{self.apply.__qualname__}")
        cargs.event_data = self
        self.on_applying(cargs)
        if cargs.cancel:
            return
        _Events().trigger(FormatNamedEvent.STYLE_APPLYING, cargs)
        if cargs.cancel:
            return

        mLo.Lo.dispatch_cmd("BackgroundColor", mProps.Props.make_props(BackgroundColor=self.prop_color))
        mLo.Lo.dispatch_cmd("Escape")

        eargs = EventArgs.from_args(cargs)
        self.on_applied(eargs)
        _Events().trigger(FormatNamedEvent.STYLE_APPLIED, eargs)

    # endregion apply()
    def dispatch_reset(self) -> None:
        """
        Resets the cursor at is current position/selection to remove any Fill Color Formatting.

        Returns:
            None:
        """
        mLo.Lo.dispatch_cmd("BackgroundColor", mProps.Props.make_props(BackgroundColor=-1))
        mLo.Lo.dispatch_cmd("Escape")

    def backup(self, obj: object) -> None:
        """Overrides, No actions are taken"""
        pass

    def restore(self, obj: object, clear: bool = False) -> None:
        """Overrides, No actions are taken"""
        self.dispatch_reset()

    # region set styles

    # endregion set styles
    @property
    def prop_has_backup(self) -> bool:
        """Gets If instantance has backup data. Overrides, Returns ``False``."""
        return True

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        return FormatKind.PARA

    @property
    def prop_color(self) -> Color:
        """Gets/Sets color"""
        return self._get("FillColor")

    @prop_color.setter
    def prop_color(self, value: Color):
        if value >= 0:
            self._set("FillColor", value)
            self._set("FillBackground", False)
        else:
            self._set("FillColor", -1)
            self._set("FillBackground", True)

    @static_prop
    def empty() -> FillColor:  # type: ignore[misc]
        """Gets FillColor empty. Static Property."""
        if FillColor._EMPTY is None:
            FillColor._EMPTY = FillColor()
        return FillColor._EMPTY
