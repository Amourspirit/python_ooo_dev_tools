from __future__ import annotations
from typing import TYPE_CHECKING, Type
from ooodev.units.unit_convert import UnitLength
from ooodev.units.unit_cm import UnitCM
from ooodev.units.unit_inch import UnitInch
from ooodev.units.unit_inch10 import UnitInch10
from ooodev.units.unit_inch100 import UnitInch100
from ooodev.units.unit_inch1000 import UnitInch1000
from ooodev.units.unit_mm import UnitMM
from ooodev.units.unit_mm10 import UnitMM10
from ooodev.units.unit_mm100 import UnitMM100
from ooodev.units.unit_pt import UnitPT
from ooodev.units.unit_px import UnitPX

if TYPE_CHECKING:
    from ooodev.units.unit_obj import UnitT


def get_unit_type(unit_length: UnitLength) -> Type[UnitT]:
    """
    Gets the unit type.

    Args:
        unit_length (UnitLength): Unit Length.

    Raises:
        ValueError: If unknown unit length or there is not type to match the unit length.

    Returns:
        Type[UnitT]: Unit Type.

    Example:
        .. code-block:: python

            from ooodev.units.unit_factory import get_unit_type
            from ooodev.units UnitLength
            unit_mm100_type = get_unit_type(UnitLength.MM100)
    .. versionadded:: 0.34.1
    """
    if unit_length == UnitLength.CM:
        return UnitCM
    elif unit_length == UnitLength.IN:
        return UnitInch
    elif unit_length == UnitLength.IN10:
        return UnitInch10
    elif unit_length == UnitLength.IN100:
        return UnitInch100
    elif unit_length == UnitLength.IN1000:
        return UnitInch1000
    elif unit_length == UnitLength.MM:
        return UnitMM
    elif unit_length == UnitLength.MM10:
        return UnitMM10
    elif unit_length == UnitLength.MM100:
        return UnitMM100
    elif unit_length == UnitLength.PT:
        return UnitPT
    elif unit_length == UnitLength.PX:
        return UnitPX
    else:
        raise ValueError(f"Unknown unit length: {unit_length}")


def get_unit(unit_length: UnitLength, value: int | float) -> UnitT:
    """
    Gets the unit.

    Args:
        unit_length (UnitLength): Unit Length.
        value (int | float): Value.

    Returns:
        UnitT: Unit.

    Example:
        .. code-block:: python

            >>> from ooodev.units.unit_factory import get_unit
            >>> from ooodev.units UnitLength
            >>> unit_mm100 = get_unit(UnitLength.MM100, 500)
            >>> print(unit_mm100)
            UnitMM100(value=500)

    .. versionadded:: 0.34.1
    """
    t = get_unit_type(unit_length)
    return t.from_unit_val(value)
