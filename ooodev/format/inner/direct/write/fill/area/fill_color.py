"""
Module for Fill Properties Fill Color.

.. versionadded:: 0.9.0
"""

# pylint: disable=unexpected-keyword-arg
# pylint: disable=invalid-name

from __future__ import annotations
from typing import Any, Tuple, TypeVar

from ooodev.format.inner.common.abstract.abstract_fill_color import AbstractColor
from ooodev.format.inner.common.props.fill_color_props import FillColorProps
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.loader import lo as mLo
from ooodev.utils import color as mColor


_TFillColor = TypeVar("_TFillColor", bound="FillColor")


class FillColor(AbstractColor):
    """
    Class for Fill Properties Fill Color.

    .. seealso::

        - :ref:`help_writer_format_direct_shape_color`

    .. versionadded:: 0.9.0
    """

    def __init__(self, color: mColor.Color = mColor.StandardColor.AUTO_COLOR) -> None:
        """
        Constructor

        Args:
            color (:py:data:`~.utils.color.Color`, optional): FillColor Color.

        Returns:
            None:

        See Also:
            :ref:`help_writer_format_direct_shape_color`
        """
        super().__init__(color=color)

    def _container_get_service_name(self) -> str:
        return super()._container_get_service_name()

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

    def _is_valid_obj(self, obj: Any) -> bool:
        if super()._is_valid_obj(obj):
            return True
        return bool(mLo.Lo.is_uno_interfaces(obj, "com.sun.star.beans.XPropertySet"))

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
            self._empty_props = self.__class__(_cattribs=self._get_internal_cattribs())  # type: ignore
            self._empty_props._is_default_inst = True  # pylint: disable=protected-access
        return self._empty_props
