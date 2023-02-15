"""
Module for managing character border side.

.. versionadded:: 0.9.0
"""
# region imports
from __future__ import annotations
from typing import Tuple

import uno
from ...common.abstract.abstract_sides import AbstractSides
from ...common.props.border_props import BorderProps
from ....kind.format_kind import FormatKind
from ...structs.side import Side as Side, LineSize as LineSize, SideFlags as SideFlags

from ooo.dyn.table.border_line_style import BorderLineStyleEnum as BorderLineStyleEnum

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
        return ("com.sun.star.style.ParagraphProperties", "com.sun.star.style.ParagraphStyle")

    # endregion methods

    # region Properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        return FormatKind.PARA | FormatKind.STATIC

    @property
    def _props(self) -> BorderProps:
        try:
            return self.__border_properties
        except AttributeError:
            self.__border_properties = BorderProps(
                left="LeftBorder", top="TopBorder", right="RightBorder", bottom="BottomBorder"
            )
        return self.__border_properties

    # endregion Properties
