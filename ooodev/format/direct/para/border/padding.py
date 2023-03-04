"""
Modele for managing paragraph padding.

.. versionadded:: 0.9.0
"""
from __future__ import annotations
from typing import Tuple

from .....meta.static_prop import static_prop
from ....kind.format_kind import FormatKind
from ...common.abstract.abstract_padding import AbstractPadding
from ...common.props.border_props import BorderProps


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
                "com.sun.star.style.ParagraphProperties",
                "com.sun.star.style.ParagraphStyle",
                "com.sun.star.style.PageStyle",
                "com.sun.star.text.TextFrame",
                "com.sun.star.text.TextGraphicObject",
                "com.sun.star.text.BaseFrame",
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
                left="LeftBorderDistance",
                top="TopBorderDistance",
                right="RightBorderDistance",
                bottom="BottomBorderDistance",
            )
        return self._props_internal_attributes

    @static_prop
    def default() -> Padding:  # type: ignore[misc]
        """Gets BorderPadding default. Static Property."""
        try:
            return Padding._DEFAULT_INST
        except AttributeError:
            inst = Padding()
            inst._set(inst._props.bottom, 0)
            inst._set(inst._props.left, 0)
            inst._set(inst._props.right, 0)
            inst._set(inst._props.top, 0)
            inst._is_default_inst = True
            Padding._DEFAULT_INST = inst
        return Padding._DEFAULT_INST

    # endregion properties
