import os
import pytest
from pathlib import Path
from typing import cast

# from ooodev.office.write import Write
if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.loader.lo import Lo
from ooodev.exceptions import ex
from ooodev.units.convert.converter import Converter
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


# Assuming UnitLengthKind is an enum with members like METERS, CENTIMETERS, etc.
@pytest.fixture(scope="module")
def the_converter() -> Converter:
    return Converter()


@pytest.mark.parametrize(
    "test_id, num, fm, to, expected",
    [
        ("happy-1", 1, UnitLengthKind.METER, UnitLengthKind.CM, 100.0),
        ("happy-2", 100, UnitLengthKind.CM, UnitLengthKind.METER, 1.0),
        ("happy-3", 1, UnitLengthKind.METER_KILO, UnitLengthKind.METER, 1000.0),
        # Add more happy path cases as needed
    ],
)
def test_convert_length_happy_path(test_id, num, fm, to, expected, loader, the_converter: Converter):
    # Act
    result = the_converter.convert_length(num, fm, to)

    # Assert
    assert result == expected, f"Failed test ID: {test_id}"


@pytest.mark.parametrize(
    "test_id, num, fm, to, expected",
    [
        ("nice-1", 1, UnitAreaKind.ARE, UnitAreaKind.ARE_CENTI, 100.0),
        ("nice-2", 45, UnitAreaKind.ARE_KILO, UnitAreaKind.ARE, 45000.0),
        ("nice-3", 1, UnitAreaKind.ARE_KILO, UnitAreaKind.ARE, 1000.0),
        # Add more happy path cases as needed
    ],
)
def test_convert_area_nice(test_id, num, fm, to, expected, loader, the_converter: Converter):
    # Act
    result = the_converter.convert_area(num, fm, to)

    # Assert
    assert result == expected, f"Failed test ID: {test_id}"


def test_convert_all_area_units(loader, the_converter: Converter):
    # Arrange

    vals = set()
    for val in UnitAreaKind:
        vals.add(val.value)
    if "ar" in vals:
        vals.remove("ar")

    for val in vals:
        # Act
        result = the_converter.convert(1, frm="ar", to=val)
        # Assert
        assert result is not None


def test_convert_all_length_units(loader, the_converter: Converter):
    # Arrange

    vals = set()
    for val in UnitLengthKind:
        vals.add(val.value)
    if "m" in vals:
        vals.remove("m")

    for val in vals:
        # Act
        result = the_converter.convert(1, frm="m", to=val)
        # Assert
        assert result is not None


def test_convert_all_energy_units(loader, the_converter: Converter):
    # Arrange

    vals = set()
    for val in UnitEnergyKind:
        vals.add(val.value)
    if "e" in vals:
        vals.remove("e")

    for val in vals:
        # Act
        result = the_converter.convert(1, frm="e", to=val)
        # Assert
        assert result is not None


def test_convert_all_flux_density_units(loader, the_converter: Converter):
    # Arrange

    vals = set()
    for val in UnitFluxDensityKind:
        vals.add(val.value)
    if "ga" in vals:
        vals.remove("ga")

    for val in vals:
        # Act
        result = the_converter.convert(1, frm="ga", to=val)
        # Assert
        assert result is not None


def test_convert_all_force_units(loader, the_converter: Converter):
    # Arrange

    vals = set()
    for val in UnitForceKind:
        vals.add(val.value)
    if "N" in vals:
        vals.remove("N")

    for val in vals:
        # Act
        result = the_converter.convert(1, frm="N", to=val)
        # Assert
        assert result is not None


def test_convert_all_info_units(loader, the_converter: Converter):
    # Arrange

    vals = set()
    for val in UnitInfoKind:
        vals.add(val.value)
    if "bit" in vals:
        vals.remove("bit")

    for val in vals:
        # Act
        result = the_converter.convert(1, frm="bit", to=val)
        # Assert
        assert result is not None


def test_convert_all_weight_units(loader, the_converter: Converter):
    # Arrange

    vals = set()
    for val in UnitWeightKind:
        vals.add(val.value)
    if "g" in vals:
        vals.remove("g")

    for val in vals:
        # Act
        result = the_converter.convert(1, frm="g", to=val)
        # Assert
        assert result is not None


def test_convert_all_power_units(loader, the_converter: Converter):
    # Arrange

    vals = set()
    for val in UnitPowerKind:
        vals.add(val.value)
    if "W" in vals:
        vals.remove("W")

    for val in vals:
        # Act
        result = the_converter.convert(1, frm="W", to=val)
        # Assert
        assert result is not None


def test_convert_all_pressure_units(loader, the_converter: Converter):
    # Arrange

    vals = set()
    for val in UnitPressureKind:
        vals.add(val.value)
    if "psi" in vals:
        vals.remove("psi")

    for val in vals:
        # Act
        result = the_converter.convert(1, frm="psi", to=val)
        # Assert
        assert result is not None


def test_convert_all_speed_units(loader, the_converter: Converter):
    # Arrange

    vals = set()
    for val in UnitSpeedKind:
        vals.add(val.value)
    if "mph" in vals:
        vals.remove("mph")

    for val in vals:
        # Act
        result = the_converter.convert(1, frm="mph", to=val)
        # Assert
        assert result is not None


def test_convert_all_temp_units(loader, the_converter: Converter):
    # Arrange

    vals = set()
    for val in UnitTempKind:
        vals.add(val.value)
    if "K" in vals:
        vals.remove("K")

    for val in vals:
        # Act
        result = the_converter.convert(1, frm="K", to=val)
        # Assert
        assert result is not None


def test_convert_all_time_units(loader, the_converter: Converter):
    # Arrange

    vals = set()
    for val in UnitTimeKind:
        vals.add(val.value)
    if "sec" in vals:
        vals.remove("sec")

    for val in vals:
        # Act
        result = the_converter.convert(1, frm="sec", to=val)
        # Assert
        assert result is not None


def test_convert_all_volume_units(loader, the_converter: Converter):
    # Arrange

    vals = set()
    for val in UnitVolumeKind:
        vals.add(val.value)
    if "oz" in vals:
        vals.remove("oz")

    for val in vals:
        # Act
        result = the_converter.convert(1, frm="oz", to=val)
        # Assert
        assert result is not None
