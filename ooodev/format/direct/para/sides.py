"""
Module for managing character border side.

.. versionadded:: 0.9.0
"""
# region imports
from __future__ import annotations
from typing import cast

import uno
from ..common.abstract_sides import AbstractSides, BorderProps
from ....exceptions import ex as mEx
from ...kind.format_kind import FormatKind
from ..structs.side import Side as Side, BorderLineStyleEnum as BorderLineStyleEnum

from ooo.dyn.table.border_line2 import BorderLine2

# endregion imports


class Sides(AbstractSides):
    """
    Paragraph Border.

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together Sides properties.

    .. versionadded:: 0.9.0
    """

    # region methods

    @staticmethod
    def from_obj(obj: object) -> Sides:
        """
        Gets instance from object properties

        Args:
            obj (object): UNO object that has supports ``com.sun.star.style.CharacterProperties`` service.

        Raises:
            NotSupportedServiceError: If ``obj`` does not support ``com.sun.star.style.CharacterProperties`` service.
            PropertyNotFoundError: If ``obj`` does not have ``TableBorder2`` property.

        Returns:
            BorderTable: Border Table.
        """
        bc = Sides()
        if not bc._is_valid_obj(obj):
            raise mEx.NotSupportedServiceError(bc._supported_services()[0])

        empty = BorderLine2()
        for attr in bc.__border_properties:
            b2 = cast(BorderLine2, getattr(obj, attr, empty))
            side = Side.from_border2(b2)
            bc._set(attr, side)
        return bc

    # endregion methods

    # region Properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        return FormatKind.PARA | FormatKind.STATIC

    @property
    def _border(self) -> BorderProps:
        try:
            return self.__border_properties
        except AttributeError:
            self.__border_properties = BorderProps(
                left="LeftBorder", top="TopBorder", right="RightBorder", bottom="BottomBorder"
            )
        return self.__border_properties

    # endregion Properties
