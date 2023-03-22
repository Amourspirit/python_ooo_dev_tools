"""
Module for Fill Properties Fill Color.

.. versionadded:: 0.9.0
"""
from __future__ import annotations
from typing import Tuple, TypeVar

from ooodev.utils import lo as mLo
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.common.abstract.abstract_fill_color import AbstractColor
from ooodev.format.inner.common.props.fill_color_props import FillColorProps


_TFillColor = TypeVar(name="_TFillColor", bound="FillColor")


class FillColor(AbstractColor):
    """
    Class for Fill Properties Fill Color.

    .. versionadded:: 0.9.0
    """

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = (
                "com.sun.star.beans.PropertySet",
                "com.sun.star.chart2.PageBackground",
                "com.sun.star.drawing.FillProperties",
                "com.sun.star.style.Style",
                "com.sun.star.text.BaseFrame",
                "com.sun.star.text.TextEmbeddedObject",
                "com.sun.star.text.TextFrame",
                "com.sun.star.text.TextGraphicObject",
            )
        return self._supported_services_values

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
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.TXT_CONTENT
        return self._format_kind_prop

    @property
    def _props(self) -> FillColorProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = FillColorProps(color="FillColor", style="FillStyle", bg="FillBackground")
        return self._props_internal_attributes

    @property
    def empty(self: _TFillColor) -> _TFillColor:  # type: ignore[misc]
        """Gets FillColor empty."""
        try:
            return self._empty_props
        except AttributeError:
            self._empty_props = self.__class__(_cattribs=self._get_internal_cattribs())
            self._empty_props._is_default_inst = True
        return self._empty_props
