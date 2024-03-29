from __future__ import annotations
from enum import Enum


class UnitAreaKind(Enum):
    """Area units"""

    ARMSTRONG_SQ = "ang2"
    """square Armstrong"""
    ACRE_INTERNATIONAL = "uk_acre"
    """acre (international)"""
    ACRE_US = "us_acre"
    """acre (US)"""
    ARE_YOTTA = "Yar"
    """yotta Are 10/24"""
    ARE_ZETTA = "Zar"
    """zetta Are 10/21"""
    ARE_EXA = "Ear"
    """exa 10/18"""
    ARE_PETA = "Par"
    """peta Are 10/15"""
    ARE_TERA = "Tar"
    """tera Are 10/12"""
    ARE_GIGA = "Gar"
    """giga Are 10/9"""
    ARE_MEGA = "Mar"
    """mega Are 10/6"""
    ARE_KILO = "kar"
    """kilo Are 10/3"""
    ARE_HECTO = "har"
    """hecto Are 10/2"""
    ARE_DECA = "ear"
    """deca Are 10/1"""
    ARE = "ar"
    """are"""
    ARE_DECI = "dar"
    """deci Are 10/-1"""
    ARE_CENTI = "car"
    """centi Are 10/-2"""
    ARE_MILLI = "mar"
    """milli Are 10/-3"""
    ARE_MICRO = "uar"
    """micro Are 10/-6"""
    ARE_NANO = "nar"
    """nano Are 10/-9"""
    ARE_PICO = "par"
    """pico Are 10/-12"""
    ARE_FEMTO = "far"
    """femto Are 10/-15"""
    ARE_ATTO = "aar"
    """atto Are 10/-18"""
    ARE_ZEPTO = "zar"
    """zepto Are 10/-21"""
    ARE_YOCTO = "yar"
    """yocto Are 10/-24"""
    FOOT_SQ = "ft2"
    """square foot"""
    HECTARE = "ha"
    """hectare"""
    INCH_SQ = "in2"
    """square inch"""
    LIGHT_YEAR_SQ = "ly2"
    """square light year"""
    MILE_INTERNATIONAL_SQ = "mi2"
    """square mile (international)"""
    MORGEN = "Morgen"
    """Morgen"""
    MILE_NAUTICAL_SQ = "Nmi2"
    """square mile (nautical)"""
    PICA_POINT_SQ = "Pica2"
    """square pica point"""
    YARD_SQ = "yd2"
    """square yard"""

    def __str__(self) -> str:
        return self.value
