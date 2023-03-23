"""
Module for managing character border side.

.. versionadded:: 0.9.0
"""
# region imports
from __future__ import annotations
from typing import Tuple

from ooodev.format.inner.common.abstract.abstract_sides import AbstractSides
from ooodev.format.inner.common.props.border_props import BorderProps
from ooodev.format.inner.kind.format_kind import FormatKind

# endregion imports


class Sides(AbstractSides):
    """
    Paragraph Border.

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together Sides properties.

    .. versionadded:: 0.9.0
    """

    # region methods
    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = (
                "com.sun.star.style.PageStyle",
                "com.sun.star.style.ParagraphProperties",
                "com.sun.star.style.ParagraphStyle",
                "com.sun.star.text.BaseFrame",
                "com.sun.star.text.TextEmbeddedObject",
                "com.sun.star.text.TextFrame",
                "com.sun.star.text.TextGraphicObject",
            )
        return self._supported_services_values

    # endregion methods

    # region Properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.PARA | FormatKind.STATIC
        return self._format_kind_prop

    @property
    def _props(self) -> BorderProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = BorderProps(
                left="LeftBorder", top="TopBorder", right="RightBorder", bottom="BottomBorder"
            )
        return self._props_internal_attributes

    # endregion Properties
