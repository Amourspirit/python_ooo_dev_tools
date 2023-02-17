"""
Module for Fill Properties Fill Color.

.. versionadded:: 0.9.0
"""
from __future__ import annotations
from typing import Tuple

from .....meta.static_prop import static_prop
from .....utils import lo as mLo
from ....kind.format_kind import FormatKind
from ...common.abstract.abstract_fill_color import AbstractColor
from ...common.props.fill_color_props import FillColorProps

from ooo.dyn.drawing.fill_style import FillStyle as FillStyle


class FillColor(AbstractColor):
    """
    Class for Fill Properties Fill Color.

    .. versionadded:: 0.9.0
    """

    def _supported_services(self) -> Tuple[str, ...]:
        return (
            "com.sun.star.drawing.FillProperties",
            "com.sun.star.beans.PropertySet",
            "com.sun.star.chart2.PageBackground",
        )

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
    def _props(self) -> FillColorProps:
        try:
            return self._props_fill_color
        except AttributeError:
            self._props_fill_color = FillColorProps(color="FillColor", style="FillStyle", bg="FillBackground")
        return self._props_fill_color

    @static_prop
    def empty() -> FillColor:  # type: ignore[misc]
        """Gets FillColor empty. Static Property."""
        try:
            return FillColor._EMPTY_PROPS
        except AttributeError:
            FillColor._EMPTY_PROPS = FillColor()
            FillColor._EMPTY_PROPS._is_default_inst = True
        return FillColor._EMPTY_PROPS
