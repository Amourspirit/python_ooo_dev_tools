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
from ...structs.side import Side as Side, BorderLineStyleEnum as BorderLineStyleEnum

# endregion imports


class Sides(AbstractSides):
    """
    Character Border.

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
                "com.sun.star.style.CharacterProperties",
                "com.sun.star.style.CharacterStyle",
            )
        return self._supported_services_values

    # endregion methods

    # region Properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._fromat_kind_prop
        except AttributeError:
            self._fromat_kind_prop = FormatKind.CHAR
        return self._fromat_kind_prop

    @property
    def _props(self) -> BorderProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = BorderProps(
                left="CharLeftBorder", top="CharTopBorder", right="CharRightBorder", bottom="CharBottomBorder"
            )
        return self._props_internal_attributes

    # endregion Properties
