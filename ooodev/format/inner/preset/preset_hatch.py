from __future__ import annotations
from typing import Dict, Any
from enum import Enum
import uno
from ooo.dyn.drawing.hatch_style import HatchStyle as HatchStyle
from ooodev.utils.color import StandardColor


# region Enum
class PresetHatchKind(Enum):
    BLACK_0_DEGREES = "Black 0 Degrees"
    """Black 0 Degrees"""
    BLACK_90_DEGREES = "Black 90 Degrees"
    """Black 00 Degrees"""
    BLACK_180_DEGREES_CROSSED = "Black 180 Degrees Crossed"
    """Black 180 Degrees Crossed"""
    BLUE_45_DEGREES = "Blue 45 Degrees"
    """Blue 45 Degrees"""
    BLUE_45_DEGREES_NEG = "Blue -45 Degrees"
    """Blue -45 Degrees"""
    BLUE_45_DEGREES_CROSSED = "Blue 45 Degrees Crossed"
    """Blue 45 Degrees Crossed"""
    GREEN_30_DEGREES = "Green 30 Degrees"
    """Green 30 Degrees"""
    GREEN_60_DEGREES = "Green 60 Degrees"
    """Green 60 Degrees"""
    GREEN_90_DEGREES_TRIPLE = "Green 90 Degrees Triple"
    """Green 90 Degrees Triple"""
    RED_45_DEGREES = "Red 45 Degrees"
    """Red 45 Degrees"""
    RED_90_DEGREES_CROSSED = "Red 90 Degrees Crossed"
    """Red 90 Degrees Crossed"""
    RED_45_DEGREES_NEG_TRIPLE = "Red -45 Degrees Triple"
    """Red -45 Degrees Triple"""
    YELLOW_45_DEGREES = "Yellow 45 Degrees"
    """Yellow 45 Degrees"""
    YELLOW_45_DEGREES_CROSSED = "Yellow 45 Degrees Crossed"
    """Yellow 45 Degrees Crossed"""
    YELLOW_45_DEGREES_TRIPLE = "Yellow 45 Degrees Triple"
    """Yellow 45 Degrees Triple"""

    def __str__(self) -> str:
        return self.value

    @staticmethod
    def is_preset(name: str) -> bool:
        """
        Gets if name is a preset name.

        Args:
            name (str): Name such as ``Red 45 Degrees``.

        Returns:
            bool: ``True`` if preset name; Otherwise, ``False``.
        """
        try:
            p_name = PresetHatchKind._preset_names
        except AttributeError:
            attrs = [getattr(PresetHatchKind, x).value for x in dir(PresetHatchKind) if x.isupper()]
            PresetHatchKind._preset_names = tuple(attrs)
            p_name = PresetHatchKind._preset_names

        return name in p_name


# endregion Enum


# region Fill Preset Dictionaries
def fill_black_0_degrees() -> Dict[str, Any]:
    """Black 0 Degrees"""
    return {
        "style": HatchStyle.SINGLE,
        "color": StandardColor.BLACK,
        "space": 1.02,
        "angle": 0,
        "bg_color": -1,
    }


def fill_black_90_degrees() -> Dict[str, Any]:
    """Black 90 Degrees"""
    return {
        "style": HatchStyle.SINGLE,
        "color": StandardColor.BLACK,
        "space": 1.02,
        "angle": 90,
        "bg_color": -1,
    }


def fill_black_180_degrees_crossed() -> Dict[str, Any]:
    """Black 180 Degrees Crossed"""
    return {
        "style": HatchStyle.DOUBLE,
        "color": StandardColor.BLACK,
        "space": 1.02,
        "angle": 180,
        "bg_color": -1,
    }


def fill_blue_45_degrees() -> Dict[str, Any]:
    """Blue 45 Degrees"""
    return {
        "style": HatchStyle.SINGLE,
        "color": StandardColor.BLUE,
        "space": 2.03,
        "angle": 45,
        "bg_color": -1,
    }


def fill_blue_45_degrees_neg() -> Dict[str, Any]:
    """Blue -45 Degrees"""
    return {
        "style": HatchStyle.SINGLE,
        "color": StandardColor.BLUE,
        "space": 2.03,
        "angle": 315,
        "bg_color": -1,
    }


def fill_blue_45_degrees_crossed() -> Dict[str, Any]:
    """Blue 45 Degrees Crossed"""
    return {
        "style": HatchStyle.DOUBLE,
        "color": StandardColor.BLUE,
        "space": 2.03,
        "angle": 45,
        "bg_color": -1,
    }


def fill_green_30_degrees() -> Dict[str, Any]:
    """Green 30 Degrees"""
    return {
        "style": HatchStyle.SINGLE,
        "color": StandardColor.GREEN,
        "space": 2.03,
        "angle": 30,
        "bg_color": -1,
    }


def fill_green_60_degrees() -> Dict[str, Any]:
    """Green 60 Degrees"""
    return {
        "style": HatchStyle.SINGLE,
        "color": StandardColor.GREEN,
        "space": 2.03,
        "angle": 60,
        "bg_color": -1,
    }


def fill_green_90_degrees_triple() -> Dict[str, Any]:
    """Green 90 Degrees Triple"""
    return {
        "style": HatchStyle.TRIPLE,
        "color": StandardColor.GREEN,
        "space": 2.03,
        "angle": 90,
        "bg_color": -1,
    }


def fill_red_45_degrees() -> Dict[str, Any]:
    """Red 45 Degrees"""
    return {
        "style": HatchStyle.SINGLE,
        "color": StandardColor.RED,
        "space": 3.05,
        "angle": 45,
        "bg_color": -1,
    }


def fill_red_90_degrees_crossed() -> Dict[str, Any]:
    """Red 90 Degrees Crossed"""
    return {
        "style": HatchStyle.DOUBLE,
        "color": StandardColor.RED,
        "space": 3.05,
        "angle": 90,
        "bg_color": -1,
    }


def fill_red_45_degrees_neg_triple() -> Dict[str, Any]:
    """Red -45 Degrees Triple"""
    return {
        "style": HatchStyle.TRIPLE,
        "color": StandardColor.RED,
        "space": 3.05,
        "angle": 135,
        "bg_color": -1,
    }


def fill_yellow_45_degrees() -> Dict[str, Any]:
    """Yellow 45 Degrees"""
    return {
        "style": HatchStyle.SINGLE,
        "color": StandardColor.GOLD,
        "space": 4.10,
        "angle": 45,
        "bg_color": -1,
    }


def fill_yellow_45_degrees_crossed() -> Dict[str, Any]:
    """Yellow 45 Degrees Crossed"""
    return {
        "style": HatchStyle.DOUBLE,
        "color": StandardColor.GOLD,
        "space": 4.10,
        "angle": 45,
        "bg_color": -1,
    }


def fill_yellow_45_degrees_triple() -> Dict[str, Any]:
    """Yellow 45 Degrees Triple"""
    return {
        "style": HatchStyle.TRIPLE,
        "color": StandardColor.GOLD,
        "space": 4.10,
        "angle": 45,
        "bg_color": -1,
    }


# endregion Fill Preset Dictionaries


# region Get Preset
def get_preset(kind: PresetHatchKind) -> Dict[str, Any]:
    """
    Gets preset

    Returns:
        PresetHatchKind: Preset Kind
    """
    if kind == PresetHatchKind.BLACK_0_DEGREES:
        return fill_black_0_degrees()
    if kind == PresetHatchKind.BLACK_90_DEGREES:
        return fill_black_90_degrees()
    if kind == PresetHatchKind.BLACK_180_DEGREES_CROSSED:
        return fill_black_180_degrees_crossed()
    if kind == PresetHatchKind.BLUE_45_DEGREES:
        return fill_blue_45_degrees()
    if kind == PresetHatchKind.BLUE_45_DEGREES_NEG:
        return fill_blue_45_degrees_neg()
    if kind == PresetHatchKind.BLUE_45_DEGREES_CROSSED:
        return fill_blue_45_degrees_crossed()
    if kind == PresetHatchKind.GREEN_30_DEGREES:
        return fill_green_30_degrees()
    if kind == PresetHatchKind.GREEN_60_DEGREES:
        return fill_green_60_degrees()
    if kind == PresetHatchKind.GREEN_90_DEGREES_TRIPLE:
        return fill_green_90_degrees_triple()
    if kind == PresetHatchKind.RED_45_DEGREES:
        return fill_red_45_degrees()
    if kind == PresetHatchKind.RED_90_DEGREES_CROSSED:
        return fill_red_90_degrees_crossed()
    if kind == PresetHatchKind.RED_45_DEGREES_NEG_TRIPLE:
        return fill_red_45_degrees_neg_triple()
    if kind == PresetHatchKind.YELLOW_45_DEGREES:
        return fill_yellow_45_degrees()
    if kind == PresetHatchKind.YELLOW_45_DEGREES_CROSSED:
        return fill_yellow_45_degrees_crossed()
    return fill_yellow_45_degrees_triple()


# endregion Get Preset
