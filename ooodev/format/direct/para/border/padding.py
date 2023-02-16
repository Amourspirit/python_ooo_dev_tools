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
        return (
            "com.sun.star.style.ParagraphProperties",
            "com.sun.star.style.ParagraphStyle",
            "com.sun.star.style.PageStyle",
        )

    # endregion methods

    # region properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        return FormatKind.PARA

    @property
    def _props(self) -> BorderProps:
        try:
            return self.__border_properties
        except AttributeError:
            self.__border_properties = BorderProps(
                left="LeftBorderDistance",
                top="TopBorderDistance",
                right="RightBorderDistance",
                bottom="BottomBorderDistance",
            )
        return self.__border_properties

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
