from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno
from ooodev.adapter.sheet.function_access_comp import FunctionAccessComp
from ooodev.loader import lo as mLo
from ooodev.exceptions import ex as mEx


if TYPE_CHECKING:
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.units.convert.unit_area_kind import UnitAreaKind
    from ooodev.units.convert.unit_length_kind import UnitLengthKind
    from ooodev.units.convert.unit_energy_kind import UnitEnergyKind
    from ooodev.units.convert.unit_flux_density_kind import UnitFluxDensityKind
    from ooodev.units.convert.unit_force_kind import UnitForceKind
    from ooodev.units.convert.unit_info_kind import UnitInfoKind
    from ooodev.units.convert.unit_weight_kind import UnitWeightKind
    from ooodev.units.convert.unit_power_kind import UnitPowerKind
    from ooodev.units.convert.unit_pressure_kind import UnitPressureKind
    from ooodev.units.convert.unit_speed_kind import UnitSpeedKind
    from ooodev.units.convert.unit_temp_kind import UnitTempKind
    from ooodev.units.convert.unit_time_kind import UnitTimeKind
    from ooodev.units.convert.unit_volume_kind import UnitVolumeKind


class Converter(FunctionAccessComp):
    """
    Converter Class.

    Converts from one unit to another. The underlying function is ``CONVERT``.

    All Conversion units are available as Enums. For example, ``UnitAreaKind``, ``UnitLengthKind``, etc.

    Note that ``CONVERT`` is the same function as in Calc.

    .. seealso::
        `Calc CONVERT function <https://help.libreoffice.org/latest/en-US/text/scalc/01/func_convert.html?&DbPAR=CALC&System=UNIX>`__

    .. versionadded:: 0.35.0
    """

    def __init__(self, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            lo_inst (LoInst): Lo Instance. This instance is used to create ``component`` is it is not provided.

        Returns:
            None:

        Note:
            ``component`` is automatically created from ``lo_inst`` if it is not provided.
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        FunctionAccessComp.__init__(self, lo_inst=lo_inst, component=None)

    def convert_area(self, num: int | float, fm: UnitAreaKind, to: UnitAreaKind) -> float:
        """
        Converts area from one unit to another.

        Args:
            fm (UnitAreaKind): Enum values, from unit.
            to (UnitAreaKind): Enum value, to unit.

        Raises:
            ConvertError: If the conversion fails.

        Returns:
            float: The converted value.

        Hint:
            - ``UnitAreaKind`` can be imported from ``ooodev.units.convert``.
        """
        return self.convert(num, str(fm), str(to))

    def convert_length(self, num: int | float, fm: UnitLengthKind, to: UnitLengthKind) -> float:
        """
        Converts length from one unit to another.

        Args:
            fm (UnitLengthKind): Enum values, from unit.
            to (UnitLengthKind): Enum value, to unit.

        Raises:
            ConvertError: If the conversion fails.

        Returns:
            float: The converted value.

        Hint:
            - ``UnitLengthKind`` can be imported from ``ooodev.units.convert``.
        """
        return self.convert(num, str(fm), str(to))

    def convert_energy(self, num: int | float, fm: UnitEnergyKind, to: UnitEnergyKind) -> float:
        """
        Converts energy from one unit to another.

        Args:
            fm (UnitEnergyKind): Enum values, from unit.
            to (UnitEnergyKind): Enum value, to unit.

        Raises:
            ConvertError: If the conversion fails.

        Returns:
            float: The converted value.

        Hint:
            - ``UnitEnergyKind`` can be imported from ``ooodev.units.convert``.
        """
        return self.convert(num, str(fm), str(to))

    def convert_flux_density(self, num: int | float, fm: UnitFluxDensityKind, to: UnitFluxDensityKind) -> float:
        """
        Converts flux density from one unit to another.

        Args:
            fm (UnitFluxDensityKind): Enum values, from unit.
            to (UnitFluxDensityKind): Enum value, to unit.

        Raises:
            ConvertError: If the conversion fails.

        Returns:
            float: The converted value.

        Hint:
            - ``UnitFluxDensityKind`` can be imported from ``ooodev.units.convert``.
        """
        return self.convert(num, str(fm), str(to))

    def convert_force(self, num: int | float, fm: UnitForceKind, to: UnitForceKind) -> float:
        """
        Converts force from one unit to another.

        Args:
            fm (UnitForceKind): Enum values, from unit.
            to (UnitForceKind): Enum value, to unit.

        Raises:
            ConvertError: If the conversion fails.

        Returns:
            float: The converted value.

        Hint:
            - ``UnitForceKind`` can be imported from ``ooodev.units.convert``.
        """
        return self.convert(num, str(fm), str(to))

    def convert_info(self, num: int | float, fm: UnitInfoKind, to: UnitInfoKind) -> float:
        """
        Converts info from one unit to another.

        Args:
            fm (UnitInfoKind): Enum values, from unit.
            to (UnitInfoKind): Enum value, to unit.

        Raises:
            ConvertError: If the conversion fails.

        Returns:
            float: The converted value.

        Hint:
            - ``UnitInfoKind`` can be imported from ``ooodev.units.convert``.
        """
        return self.convert(num, str(fm), str(to))

    def convert_weight(self, num: int | float, fm: UnitWeightKind, to: UnitWeightKind) -> float:
        """
        Converts weight from one unit to another.

        Args:
            fm (UnitWeightKind): Enum values, from unit.
            to (UnitWeightKind): Enum value, to unit.

        Raises:
            ConvertError: If the conversion fails.

        Returns:
            float: The converted value.

        Hint:
            - ``UnitWeightKind`` can be imported from ``ooodev.units.convert``.
        """
        return self.convert(num, str(fm), str(to))

    def convert_power(self, num: int | float, fm: UnitPowerKind, to: UnitPowerKind) -> float:
        """
        Converts power from one unit to another.

        Args:
            fm (UnitPowerKind): Enum values, from unit.
            to (UnitPowerKind): Enum value, to unit.

        Raises:
            ConvertError: If the conversion fails.

        Returns:
            float: The converted value.

        Hint:
            - ``UnitPowerKind`` can be imported from ``ooodev.units.convert``.
        """
        return self.convert(num, str(fm), str(to))

    def convert_pressure(self, num: int | float, fm: UnitPressureKind, to: UnitPressureKind) -> float:
        """
        Converts pressure from one unit to another.

        Args:
            fm (UnitPressureKind): Enum values, from unit.
            to (UnitPressureKind): Enum value, to unit.

        Raises:
            ConvertError: If the conversion fails.

        Returns:
            float: The converted value.

        Hint:
            - ``UnitPressureKind`` can be imported from ``ooodev.units.convert``.
        """
        return self.convert(num, str(fm), str(to))

    def convert_speed(self, num: int | float, fm: UnitSpeedKind, to: UnitSpeedKind) -> float:
        """
        Converts speed from one unit to another.

        Args:
            fm (UnitSpeedKind): Enum values, from unit.
            to (UnitSpeedKind): Enum value, to unit.

        Raises:
            ConvertError: If the conversion fails.

        Returns:
            float: The converted value.

        Hint:
            - ``UnitSpeedKind`` can be imported from ``ooodev.units.convert``.
        """
        return self.convert(num, str(fm), str(to))

    def convert_temp(self, num: int | float, fm: UnitTempKind, to: UnitTempKind) -> float:
        """
        Converts temperature from one unit to another.

        Args:
            fm (UnitTempKind): Enum values, from unit.
            to (UnitTempKind): Enum value, to unit.

        Raises:
            ConvertError: If the conversion fails.

        Returns:
            float: The converted value.

        Hint:
            - ``UnitTempKind`` can be imported from ``ooodev.units.convert``.
        """
        return self.convert(num, str(fm), str(to))

    def convert_time(self, num: int | float, fm: UnitTimeKind, to: UnitTimeKind) -> float:
        """
        Converts time from one unit to another.

        Args:
            fm (UnitTimeKind): Enum values, from unit.
            to (UnitTimeKind): Enum value, to unit.

        Raises:
            ConvertError: If the conversion fails.

        Returns:
            float: The converted value.

        Hint:
            - ``UnitTimeKind`` can be imported from ``ooodev.units.convert``.
        """
        return self.convert(num, str(fm), str(to))

    def convert_volume(self, num: int | float, fm: UnitVolumeKind, to: UnitVolumeKind) -> float:
        """
        Converts volume from one unit to another.

        Args:
            fm (UnitVolumeKind): Enum values, from unit.
            to (UnitVolumeKind): Enum value, to unit.

        Raises:
            ConvertError: If the conversion fails.

        Returns:
            float: The converted value.

        Hint:
            - ``UnitVolumeKind`` can be imported from ``ooodev.units.convert``.
        """
        return self.convert(num, str(fm), str(to))

    def convert(self, val: Any, frm: str, to: str) -> Any:
        """
        Converts a value from one unit to another.

        Args:
            val (Any): The value to be converted.
            frm (str): The unit to convert from.
            to (str): The unit to convert to.

        Raises:
            ConvertError: If the conversion fails.

        Returns:
            Any: The converted value.
        """
        try:
            return self.call_function("CONVERT", val, frm, to)
        except Exception as e:
            raise mEx.ConvertError(val, frm, to) from e
