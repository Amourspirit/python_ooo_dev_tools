from __future__ import annotations
from typing import Dict, Any
from enum import Enum
import uno
from ooo.dyn.awt.gradient_style import GradientStyle


class PresetKind(Enum):
    PASTEL_BOUQUET = 1
    PASTEL_DREAM = 2
    BLUE_TOUCH = 3
    BLANK_GRAY = 4
    SPOTTED_GRAY = 5
    LONDON_MIST = 6
    TEAL_BLUE = 7
    MIDNIGHT = 8
    DEEP_OCEAN = 9
    SUBMARINE = 10
    GREEN_GRASS = 11
    NEON_LIGHT = 12
    SUNSHINE = 13
    PRESENT = 14
    MAHOGANY = 15


def pastel_bouquet() -> Dict[str, Any]:
    """Pastel Bouquet preset"""
    return {
        "style": GradientStyle.LINEAR,
        "step_count": 0,
        "x_offset": 0,
        "y_offset": 0,
        "angle": 30,
        "border": 0,
        "start_color": 14543051,
        "start_intensity": 100,
        "end_color": 16766935,
        "end_intensity": 100,
        "name": "Pastel Bouquet",
    }


def pastel_dream() -> Dict[str, Any]:
    """Pastel Dream preset"""
    return {
        "style": GradientStyle.RECT,
        "step_count": 0,
        "x_offset": 50,
        "y_offset": 50,
        "angle": 45,
        "border": 0,
        "start_color": 16766935,
        "start_intensity": 100,
        "end_color": 11847644,
        "end_intensity": 100,
        "name": "Pastel Dream",
    }


def blue_touch() -> Dict[str, Any]:
    """Blue Touch preset"""
    return {
        "style": GradientStyle.LINEAR,
        "step_count": 0,
        "x_offset": 0,
        "y_offset": 0,
        "angle": 1,
        "border": 0,
        "start_color": 11847644,
        "start_intensity": 100,
        "end_color": 14608111,
        "end_intensity": 100,
        "name": "Blue Touch",
    }


def blank_gray() -> Dict[str, Any]:
    """Blank With Gray preset"""
    return {
        "style": GradientStyle.LINEAR,
        "step_count": 0,
        "x_offset": 0,
        "y_offset": 0,
        "angle": 90,
        "border": 75,
        "start_color": 16777215,
        "start_intensity": 100,
        "end_color": 14540253,
        "end_intensity": 100,
        "name": "Blank with Gray",
    }


def spotted_gray() -> Dict[str, Any]:
    """Spotted Gray preset"""
    return {
        "style": GradientStyle.RADIAL,
        "step_count": 0,
        "x_offset": 50,
        "y_offset": 50,
        "angle": 0,
        "border": 0,
        "start_color": 11711154,
        "start_intensity": 100,
        "end_color": 15658734,
        "end_intensity": 100,
        "name": "Spotted Gray",
    }


def london_mist() -> Dict[str, Any]:
    """London Mist preset"""
    return {
        "style": GradientStyle.LINEAR,
        "step_count": 0,
        "x_offset": 0,
        "y_offset": 0,
        "angle": 30,
        "border": 0,
        "start_color": 13421772,
        "start_intensity": 100,
        "end_color": 6710886,
        "end_intensity": 100,
        "name": "London Mist",
    }


def teal_blue() -> Dict[str, Any]:
    """Teal to Blue preset"""
    return {
        "style": GradientStyle.LINEAR,
        "step_count": 0,
        "x_offset": 0,
        "y_offset": 0,
        "angle": 30,
        "border": 0,
        "start_color": 5280650,
        "start_intensity": 100,
        "end_color": 5866416,
        "end_intensity": 100,
        "name": "Teal to Blue",
    }


def midnight() -> Dict[str, Any]:
    """Midnight preset"""
    return {
        "style": GradientStyle.LINEAR,
        "step_count": 0,
        "x_offset": 0,
        "y_offset": 0,
        "angle": 0,
        "border": 0,
        "start_color": 0,
        "start_intensity": 100,
        "end_color": 2777241,
        "end_intensity": 100,
        "name": "Midnight",
    }


def deep_ocean() -> Dict[str, Any]:
    """Deep Ocean preset"""
    return {
        "style": GradientStyle.RADIAL,
        "step_count": 0,
        "x_offset": 50,
        "y_offset": 50,
        "angle": 0,
        "border": 0,
        "start_color": 0,
        "start_intensity": 100,
        "end_color": 7512015,
        "end_intensity": 100,
        "name": "Deep Ocean",
    }


def submarine() -> Dict[str, Any]:
    """Submarine preset"""
    return {
        "style": GradientStyle.LINEAR,
        "step_count": 0,
        "x_offset": 0,
        "y_offset": 0,
        "angle": 0,
        "border": 0,
        "start_color": 14543051,
        "start_intensity": 100,
        "end_color": 11847644,
        "end_intensity": 100,
        "name": "Submarine",
    }


def green_grass() -> Dict[str, Any]:
    """Green Grass preset"""
    return {
        "style": GradientStyle.LINEAR,
        "step_count": 0,
        "x_offset": 0,
        "y_offset": 0,
        "angle": 30,
        "border": 0,
        "start_color": 16776960,
        "start_intensity": 100,
        "end_color": 8508442,
        "end_intensity": 100,
        "name": "Green Grass",
    }


def neon_light() -> Dict[str, Any]:
    """Neon Light preset"""
    return {
        "style": GradientStyle.ELLIPTICAL,
        "step_count": 0,
        "x_offset": 50,
        "y_offset": 50,
        "angle": 0,
        "border": 15,
        "start_color": 1209890,
        "start_intensity": 100,
        "end_color": 16777215,
        "end_intensity": 100,
        "name": "Neon Light",
    }


def sunshine() -> Dict[str, Any]:
    """Sunshine preset"""
    return {
        "style": GradientStyle.RADIAL,
        "step_count": 0,
        "x_offset": 66,
        "y_offset": 33,
        "angle": 0,
        "border": 33,
        "start_color": 16760576,
        "start_intensity": 100,
        "end_color": 16776960,
        "end_intensity": 100,
        "name": "Sunshine",
    }


def present() -> Dict[str, Any]:
    """Present preset"""
    return {
        "style": GradientStyle.SQUARE,
        "step_count": 0,
        "x_offset": 70,
        "y_offset": 60,
        "angle": 45,
        "border": 72,
        "start_color": 8468233,
        "start_intensity": 100,
        "end_color": 16728064,
        "end_intensity": 100,
        "name": "Present",
    }


def mahogany() -> Dict[str, Any]:
    """Mahogany preset"""
    return {
        "style": GradientStyle.SQUARE,
        "step_count": 0,
        "x_offset": 50,
        "y_offset": 50,
        "angle": 45,
        "border": 0,
        "start_color": 0,
        "start_intensity": 100,
        "end_color": 9250846,
        "end_intensity": 100,
        "name": "Mahogany",
    }


def get_preset(kind: PresetKind) -> Dict[str:Any]:
    """
    Gets preset

    Returns:
        PresetKind: Preset Kind
    """
    if kind == PresetKind.PASTEL_BOUQUET:
        return pastel_bouquet()
    if kind == PresetKind.PASTEL_DREAM:
        return pastel_dream()
    if kind == PresetKind.BLUE_TOUCH:
        return blue_touch()
    if kind == PresetKind.BLANK_GRAY:
        return blank_gray()
    if kind == PresetKind.SPOTTED_GRAY:
        return spotted_gray()
    if kind == PresetKind.LONDON_MIST:
        return london_mist()
    if kind == PresetKind.TEAL_BLUE:
        return teal_blue()
    if kind == PresetKind.MIDNIGHT:
        return midnight()
    if kind == PresetKind.DEEP_OCEAN:
        return deep_ocean()
    if kind == PresetKind.SUBMARINE:
        return submarine()
    if kind == PresetKind.GREEN_GRASS:
        return green_grass()
    if kind == PresetKind.NEON_LIGHT:
        return neon_light()
    if kind == PresetKind.SUNSHINE:
        return sunshine()
    if kind == PresetKind.PRESENT:
        return present()
    if kind == PresetKind.MAHOGANY:
        return mahogany()
