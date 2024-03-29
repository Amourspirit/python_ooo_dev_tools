from __future__ import annotations
from enum import Enum


class UnitSpeedKind(Enum):
    """Speed units"""

    KNOT_ADMIRALTY = "admkn"
    """Admiralty knot 10/0"""
    KNOT = "kn"
    """Knot (international) 10/0"""
    METER_PH_YOTTA = "Ym/h"
    """yotta Meters per hour 10/24"""
    METER_PH_ZETTA = "Zm/h"
    """zetta Meters per hour 10/21"""
    METER_PH_EXA = "Em/h"
    """exa 10/18"""
    METER_PH_PETA = "Pm/h"
    """peta Meters per hour 10/15"""
    METER_PH_TERA = "Tm/h"
    """tera Meters per hour 10/12"""
    METER_PH_GIGA = "Gm/h"
    """giga Meters per hour 10/9"""
    METER_PH_MEGA = "Mm/h"
    """mega Meters per hour 10/6"""
    METER_PH_KILO = "km/h"
    """kilo Meters per hour 10/3"""
    METER_PH_HECTO = "hm/h"
    """hecto Meters per hour 10/2"""
    METER_PH_DECA = "em/h"
    """deca Meters per hour 10/1"""
    METER_PH = "m/h"
    """Meters per hour 10/0"""
    METER_PH_DECI = "dm/h"
    """deci Meters per hour 10/-1"""
    METER_PH_CENTI = "cm/h"
    """centi Meters per hour 10/-2"""
    METER_PH_MILLI = "mm/h"
    """milli Meters per hour 10/-3"""
    METER_PH_MICRO = "um/h"
    """micro Meters per hour 10/-6"""
    METER_PH_NANO = "nm/h"
    """nano Meters per hour 10/-9"""
    METER_PH_PICO = "pm/h"
    """pico Meters per hour 10/-12"""
    METER_PH_FEMTO = "fm/h"
    """femto Meters per hour 10/-15"""
    METER_PH_ATTO = "am/h"
    """atto Meters per hour 10/-18"""
    METER_PH_ZEPTO = "zm/h"
    """zepto Meters per hour 10/-21"""
    METER_PH_YOCTO = "ym/h"
    """yocto Meters per hour 10/-24"""
    METER_PS_YOTTA = "Ym/s"
    """yotta Meters per second 10/24"""
    METER_PS_ZETTA = "Zm/s"
    """zetta Meters per second 10/21"""
    METER_PS_EXA = "Em/s"
    """exa 10/18"""
    METER_PS_PETA = "Pm/s"
    """peta Meters per second 10/15"""
    METER_PS_TERA = "Tm/s"
    """tera Meters per second 10/12"""
    METER_PS_GIGA = "Gm/s"
    """giga Meters per second 10/9"""
    METER_PS_MEGA = "Mm/s"
    """mega Meters per second 10/6"""
    METER_PS_KILO = "km/s"
    """kilo Meters per second 10/3"""
    METER_PS_HECTO = "hm/s"
    """hecto Meters per second 10/2"""
    METER_PS_DECA = "em/s"
    """deca Meters per second 10/1"""
    METER_PS = "m/s"
    """Meters per second 10/0"""
    METER_PS_DECI = "dm/s"
    """deci Meters per second 10/-1"""
    METER_PS_CENTI = "cm/s"
    """centi Meters per second 10/-2"""
    METER_PS_MILLI = "mm/s"
    """milli Meters per second 10/-3"""
    METER_PS_MICRO = "um/s"
    """micro Meters per second 10/-6"""
    METER_PS_NANO = "nm/s"
    """nano Meters per second 10/-9"""
    METER_PS_PICO = "pm/s"
    """pico Meters per second 10/-12"""
    METER_PS_FEMTO = "fm/s"
    """femto Meters per second 10/-15"""
    METER_PS_ATTO = "am/s"
    """atto Meters per second 10/-18"""
    METER_PS_ZEPTO = "zm/s"
    """zepto Meters per second 10/-21"""
    METER_PS_YOCTO = "ym/s"
    """yocto Meters per second 10/-24"""
    MILE_PH = "mph"
    """Miles per hour 10/0"""
    MPH = MILE_PH
    """Miles per hour 10/0"""

    def __str__(self) -> str:
        return self.value
