"""
Module for managing character border side.

.. versionadded:: 0.9.0
"""
# region imports
from __future__ import annotations
from typing import Tuple, Type, cast, TypeVar

import uno
from ...common.abstract_sides import AbstractSides
from .....exceptions import ex as mEx
from ....kind.format_kind import FormatKind
from ...structs.side import Side as Side, BorderLineStyleEnum as BorderLineStyleEnum

from ooo.dyn.table.border_line2 import BorderLine2

# endregion imports

_TSides = TypeVar(name="_TSides", bound="Sides")


class Sides(AbstractSides):
    """
    Character Border.

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together Sides properties.

    .. versionadded:: 0.9.0
    """

    # region methods
    def _supported_services(self) -> Tuple[str, ...]:
        return ("com.sun.star.style.CharacterProperties", "com.sun.star.style.CharacterStyle")

    @classmethod
    def from_obj(cls: Type[_TSides], obj: object) -> _TSides:
        """
        Gets instance from object properties

        Args:
            obj (object): UNO object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.
            PropertyNotFoundError: If ``obj`` does not have ``TableBorder2`` property.

        Returns:
            Sides: ``Sides`` instance that represents ``obj`` Side properties.
        """
        inst = super(Sides, cls).__new__(cls)
        inst.__init__()
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')

        empty = BorderLine2()
        for attr in inst._props:
            b2 = cast(BorderLine2, getattr(obj, attr, empty))
            side = Side.from_border2(b2)
            inst._set(attr, side)
        return inst

    # endregion methods

    # region Properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        return FormatKind.CHAR

    # endregion Properties
