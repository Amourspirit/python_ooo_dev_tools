from __future__ import annotations
from enum import Enum


class UnitPowerKind(Enum):
    """Power units"""

    HP = "HP"
    """Horsepower 10/0"""
    PS = "PS"
    """Pferdestärke or metric horsepower 10/0"""
    WATT_YOTTA = "YW"
    """yotta Watt 10/24"""
    WATT_ZETTA = "ZW"
    """zetta Watt 10/21"""
    WATT_EXA = "EW"
    """exa 10/18"""
    WATT_PETA = "PW"
    """peta Watt 10/15"""
    WATT_TERA = "TW"
    """tera Watt 10/12"""
    WATT_GIGA = "GW"
    """giga Watt 10/9"""
    WATT_MEGA = "MW"
    """mega Watt 10/6"""
    WATT_KILO = "kW"
    """kilo Watt 10/3"""
    WATT_HECTO = "hW"
    """hecto Watt 10/2"""
    WATT_DECA = "eW"
    """deca Watt 10/1"""
    WATT = "W"
    """Watt 10/0"""
    WATT_DECI = "dW"
    """deci Watt 10/-1"""
    WATT_CENTI = "cW"
    """centi Watt 10/-2"""
    WATT_MILLI = "mW"
    """milli Watt 10/-3"""
    WATT_MICRO = "uW"
    """micro Watt 10/-6"""
    WATT_NANO = "nW"
    """nano Watt 10/-9"""
    WATT_PICO = "pW"
    """pico Watt 10/-12"""
    WATT_FEMTO = "fW"
    """femto Watt 10/-15"""
    WATT_ATTO = "aW"
    """atto Watt 10/-18"""
    WATT_ZEPTO = "zW"
    """zepto Watt 10/-21"""
    WATT_YOCTO = "yW"
    """yocto Watt 10/-24"""

    def __str__(self) -> str:
        return self.value
