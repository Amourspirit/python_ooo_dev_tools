from __future__ import annotations
from typing import Tuple

from ....kind.format_kind import FormatKind
from ...common.props.border_props import BorderProps as BorderProps
from ...common.abstract.abstract_padding import AbstractPadding


class Spacing(AbstractPadding):
    """
    Frame Spacing

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together properties.

    .. versionadded:: 0.9.0
    """

    # region Overrides

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.style.Style", "com.sun.star.text.TextFrame")
        return self._supported_services_values

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.DOC | FormatKind.STYLE
        return self._format_kind_prop

    @property
    def _props(self) -> BorderProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = BorderProps(
                left="LeftMargin", top="TopMargin", right="RightMargin", bottom="BottomMargin"
            )
        return self._props_internal_attributes

    # endregion Overrides
