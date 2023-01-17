"""
Module for Shadow format (``LineSpacing``) struct.

.. versionadded:: 0.9.0
"""
# region imports
from __future__ import annotations
from typing import Dict, Tuple, overload

import uno
from ....events.event_singleton import _Events
from ....meta.static_prop import static_prop
from ....utils import props as mProps
from ...kind.format_kind import FormatKind
from ...style_base import StyleBase, EventArgs, CancelEventArgs, FormatNamedEvent

from ooo.dyn.style.line_spacing import LineSpacing as UnoLineSpacing


# endregion imports
class LineSpacing(StyleBase):
    """
    Line Spacing struct

    .. versionadded:: 0.9.0
    """

    # region init
    _EMPTY = None

    def __init__(self, mode: int = 0, height: int = 0) -> None:
        """
        Constructor

        Args:
            mode (int, optional): This value specifies the way the height is specified.
            height (int, optional): This value specifies the height in regard to Mode.

        Raises:
            ValueError: If ``color`` or ``width`` are less than zero.
        """
        if height < 0:
            raise ValueError("mode must be a positive number")
        if mode < 0:
            raise ValueError("height must be a postivie number")
        init_vals = {"Height": height, "Mode": mode}
        super().__init__(**init_vals)

    # endregion init

    # region methods

    def _supported_services(self) -> Tuple[str, ...]:
        return ()

    # region apply()

    @overload
    def apply(self, obj: object, *, keys: Dict[str, str]) -> None:
        ...

    @overload
    def apply(self, obj: object) -> None:
        ...

    def apply(self, obj: object, **kwargs) -> None:
        """
        Applies style to object

        Args:
            obj (object): Object that contains a ``LineSpacing`` property.
            keys: (Dict[str, str], optional): key map for properties.
                Can be ``spacing`` which maps to ``ParaLineSpacing`` by default.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.format_named_event.FormatNamedEvent.STYLE_APPLYING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.format_named_event.FormatNamedEvent.STYLE_APPLYED` :eventref:`src-docs-event`


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

        keys = {"spacing": "ParaLineSpacing"}
        if "keys" in kwargs:
            keys.update(kwargs["keys"])
        key = keys["spacing"]
        mProps.Props.set(obj, **{key: self.get_line_spacing()})
        eargs = EventArgs.from_args(cargs)
        self.on_applied(eargs)
        _Events().trigger(FormatNamedEvent.STYLE_APPLIED, eargs)

    # endregion apply()

    def get_line_spacing(self) -> UnoLineSpacing:
        """gets Line spacing of instance"""
        line = UnoLineSpacing()  # create the border line
        for key, val in self._dv.items():
            setattr(line, key, val)
        return line

    # endregion methods

    # region Properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        return FormatKind.STRUCT

    @property
    def prop_mode(self) -> int:
        """Gets mode value"""
        return self._get("Mode")

    @prop_mode.setter
    def prop_mode(self, value: int) -> None:
        self._set("Mode", value)

    @property
    def prop_height(self) -> int:
        """Gets the size of the shadow (in mm units)"""
        return self._get("Height")

    @prop_height.setter
    def prop_height(self, value: int) -> None:
        self._set("Height", value)

    @static_prop
    def empty() -> LineSpacing:  # type: ignore[misc]
        """Gets empty Line Spacing. Static Property."""
        if LineSpacing._EMPTY is None:
            LineSpacing._EMPTY = LineSpacing(0, 0)
        return LineSpacing._EMPTY

    # endregion Properties
