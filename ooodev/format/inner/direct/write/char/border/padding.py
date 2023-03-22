"""
Module for managing character padding.

.. versionadded:: 0.9.0
"""
# region Import
from __future__ import annotations
from typing import Tuple, TypeVar

from ooodev.format.kind.format_kind import FormatKind
from ooodev.format.inner.common.abstract.abstract_padding import AbstractPadding
from ooodev.format.inner.common.props.border_props import BorderProps

# endregion Import

_TPadding = TypeVar(name="_TPadding", bound="Padding")


class Padding(AbstractPadding):
    """
    Paragraph Border Padding

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together properties.

    .. versionadded:: 0.9.0
    """

    # region methods
    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = (
                "com.sun.star.style.CharacterProperties",
                "com.sun.star.style.CharacterStyle",
            )
        return self._supported_services_values

    # endregion methods

    # region properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.PARA
        return self._format_kind_prop

    @property
    def _props(self) -> BorderProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = BorderProps(
                left="CharLeftBorderDistance",
                top="CharTopBorderDistance",
                right="CharRightBorderDistance",
                bottom="CharBottomBorderDistance",
            )
        return self._props_internal_attributes

    @property
    def default(self: _TPadding) -> _TPadding:  # type: ignore[misc]
        """Gets BorderPadding default."""
        try:
            return self._default_inst
        except AttributeError:
            inst = self.__class__(all=0.0, _cattribs=self._get_internal_cattribs())
            inst._is_default_inst = True
            self._default_inst = inst
        return self._default_inst

    # endregion properties
