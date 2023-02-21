"""
Module for Paragraph Fill Color.

.. versionadded:: 0.9.0
"""
from __future__ import annotations
from typing import Tuple

from .....meta.static_prop import static_prop
from .....utils import lo as mLo
from .....utils import props as mProps
from ....kind.format_kind import FormatKind
from ...common.props.fill_color_props import FillColorProps
from ...common.abstract.abstract_fill_color import AbstractColor


class Color(AbstractColor):
    """
    Paragraph Fill Coloring

    .. versionadded:: 0.9.0
    """

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = (
                "com.sun.star.drawing.FillProperties",
                "com.sun.star.text.TextContent",
                "com.sun.star.style.ParagraphStyle",
                "com.sun.star.style.PageStyle",
            )
        return self._supported_services_values

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
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.PARA | FormatKind.TXT_CONTENT
        return self._format_kind_prop

    @property
    def _props(self) -> FillColorProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = FillColorProps(color="FillColor", style="FillStyle")
        return self._props_internal_attributes

    @static_prop
    def default() -> Color:  # type: ignore[misc]
        """Gets FillColor empty. Static Property."""
        try:
            return Color._DEFAULT_INST
        except AttributeError:
            Color._DEFAULT_INST = Color(-1)
            Color._DEFAULT_INST._is_default_inst = True
        return Color._DEFAULT_INST
