from __future__ import annotations
from typing import Tuple
from ...cell.background.color import Color as CellColor
from ...common.props.cell_background_color_props import CellBackgroundColorProps


class Color(CellColor):
    """
    Class for Cell Properties Back Color.

    .. versionadded:: 0.9.0
    """

    # region overrides
    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.text.TextTable",)
        return self._supported_services_values

    # endregion overrides

    # region Properties

    @property
    def _props(self) -> CellBackgroundColorProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = CellBackgroundColorProps(
                color="BackColor", is_transparent="BackTransparent"
            )
        return self._props_internal_attributes

    # endregion Properties
