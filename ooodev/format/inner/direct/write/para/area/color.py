"""
Module for Paragraph Fill Color.

.. versionadded:: 0.9.0
"""

from __future__ import annotations
from typing import Tuple

from ooodev.format.inner.common.abstract.abstract_fill_color import AbstractColor
from ooodev.format.inner.common.props.fill_color_props import FillColorProps
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.loader import lo as mLo
from ooodev.utils import color as mColor
from ooodev.utils import props as mProps
from ooodev.utils.color import StandardColor


class Color(AbstractColor):
    """
    Paragraph Fill Coloring

    .. seealso::

        - :ref:`help_writer_format_direct_para_area_color`

    .. versionadded:: 0.9.0
    """

    def __init__(self, color: mColor.Color = StandardColor.AUTO_COLOR) -> None:
        """
        Constructor

        Args:
            color (:py:data:`~.utils.color.Color`, optional): FillColor Color.

        Returns:
            None:

        See Also:

            - :ref:`help_writer_format_direct_para_area_color`
        """
        super().__init__(color)

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
