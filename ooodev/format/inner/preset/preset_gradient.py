from __future__ import annotations
from typing import Dict, Any
from enum import Enum
import uno  # pylint: disable=no-name-in-module
from ooo.dyn.awt.gradient_style import GradientStyle
from ooodev.utils.data_type.offset import Offset
from ooodev.utils.data_type.color_range import ColorRange
from ooodev.utils.data_type.intensity_range import IntensityRange
from ooodev.utils.kind import kind_helper
from ooodev.utils.color import Color

# See Also:
# MCGR - MultiColorGradients in LibreOffice - LibreOffice Conference 2023
# https://www.youtube.com/watch?v=xULTLVm2CAM
#
# Using multicolor gradients in file markup - LibreOffice Conference 2023
# https://youtu.be/X54R73Y7-Vs?si=9k5uD869PqHShr88


class PresetGradientKind(Enum):
    """Preset Gradient Kind."""

    PASTEL_BOUQUET = "Pastel Bouquet"
    """Pastel Bouquet"""
    PASTEL_DREAM = "Pastel Dream"
    """Pastel Dream"""
    BLUE_TOUCH = "Blue Touch"
    """Blue Touch"""
    BLANK_GRAY = "Blank with Gray"
    """Blank with Gray"""
    SPOTTED_GRAY = "Spotted Gray"
    """Spotted Gray"""
    LONDON_MIST = "London Mist"
    """London Mist"""
    TEAL_BLUE = "Teal to Blue"
    """Teal to Blue"""
    MIDNIGHT = "Midnight"
    """Midnight"""
    DEEP_OCEAN = "Deep Ocean"
    """Deep Ocean"""
    SUBMARINE = "Submarine"
    """Submarine"""
    GREEN_GRASS = "Green Grass"
    """Green Grass"""
    NEON_LIGHT = "Neon Light"
    """Neon Light"""
    SUNSHINE = "Sunshine"
    """Sunshine"""
    PRESENT = "Present"
    """Present"""
    MAHOGANY = "Mahogany"
    """Mahogany"""

    def __str__(self) -> str:
        return self.value

    @staticmethod
    def is_preset(name: str) -> bool:
        """
        Gets if name is a preset name.

        Args:
            name (str): Name such as ``Deep Ocean``.

        Returns:
            bool: ``True`` if preset name; Otherwise, ``False``.
        """
        try:
            p_name = PresetGradientKind._preset_names
        except AttributeError:
            attrs = [getattr(PresetGradientKind, x).value for x in dir(PresetGradientKind) if x.isupper()]
            PresetGradientKind._preset_names = tuple(attrs)
            p_name = PresetGradientKind._preset_names

        return name in p_name

    @staticmethod
    def from_str(s: str) -> PresetGradientKind:
        """
        Gets an ``PresetGradientKind`` instance from string.

        Args:
            s (str): String that represents the name of an enum Name.
                ``s`` is case insensitive and can be ``CamelCase``, ``pascal_case`` , ``snake_case``,
                ``hyphen-case``, ``normal case``.

        Raises:
            ValueError: If input string is empty.
            AttributeError: If unable to get ``PresetGradientKind`` instance.

        Returns:
            PresetGradientKind: Enum instance.
        """
        s = s.lower()
        if s == "pastel bouquet":
            return PresetGradientKind.PASTEL_BOUQUET
        elif s == "pastel dream":
            return PresetGradientKind.PASTEL_DREAM
        elif s == "blue touch":
            return PresetGradientKind.BLUE_TOUCH
        elif s == "blank with gray":
            return PresetGradientKind.BLANK_GRAY
        elif s == "spotted gray":
            return PresetGradientKind.SPOTTED_GRAY
        elif s == "london mist":
            return PresetGradientKind.LONDON_MIST
        elif s == "teal to blue":
            return PresetGradientKind.TEAL_BLUE
        elif s == "midnight":
            return PresetGradientKind.MIDNIGHT
        elif s == "deep ocean":
            return PresetGradientKind.DEEP_OCEAN
        elif s == "submarine":
            return PresetGradientKind.SUBMARINE
        elif s == "green grass":
            return PresetGradientKind.GREEN_GRASS
        elif s == "neon light":
            return PresetGradientKind.NEON_LIGHT
        elif s == "sunshine":
            return PresetGradientKind.SUNSHINE
        elif s == "present":
            return PresetGradientKind.PRESENT
        elif s == "mahogany":
            return PresetGradientKind.MAHOGANY
        else:
            return kind_helper.enum_from_string(s, PresetGradientKind)


def pastel_bouquet() -> Dict[str, Any]:
    """Pastel Bouquet preset"""
    return {
        "style": GradientStyle.LINEAR,
        "step_count": 0,
        "offset": Offset(0, 0),
        "angle": 30,
        "border": 0,
        "grad_color": ColorRange(Color(14543051), Color(16766935)),
        "grad_intensity": IntensityRange(100, 100),
        "name": PresetGradientKind.PASTEL_BOUQUET.value,
    }


def pastel_dream() -> Dict[str, Any]:
    """Pastel Dream preset"""
    return {
        "style": GradientStyle.RECT,
        "step_count": 0,
        "offset": Offset(50, 50),
        "angle": 45,
        "border": 0,
        "grad_color": ColorRange(Color(16766935), Color(11847644)),
        "grad_intensity": IntensityRange(100, 100),
        "name": PresetGradientKind.PASTEL_DREAM.value,
    }


def blue_touch() -> Dict[str, Any]:
    """Blue Touch preset"""
    return {
        "style": GradientStyle.LINEAR,
        "step_count": 0,
        "offset": Offset(0, 0),
        "angle": 1,
        "border": 0,
        "grad_color": ColorRange(Color(11847644), Color(14608111)),
        "grad_intensity": IntensityRange(100, 100),
        "name": PresetGradientKind.BLUE_TOUCH.value,
    }


def blank_gray() -> Dict[str, Any]:
    """Blank With Gray preset"""
    return {
        "style": GradientStyle.LINEAR,
        "step_count": 0,
        "offset": Offset(0, 0),
        "angle": 90,
        "border": 75,
        "grad_color": ColorRange(Color(16777215), Color(14540253)),
        "grad_intensity": IntensityRange(100, 100),
        "name": PresetGradientKind.BLANK_GRAY.value,
    }


def spotted_gray() -> Dict[str, Any]:
    """Spotted Gray preset"""
    return {
        "style": GradientStyle.RADIAL,
        "step_count": 0,
        "offset": Offset(50, 50),
        "angle": 0,
        "border": 0,
        "grad_color": ColorRange(Color(11711154), Color(15658734)),
        "grad_intensity": IntensityRange(100, 100),
        "name": PresetGradientKind.SPOTTED_GRAY.value,
    }


def london_mist() -> Dict[str, Any]:
    """London Mist preset"""
    return {
        "style": GradientStyle.LINEAR,
        "step_count": 0,
        "offset": Offset(0, 0),
        "angle": 30,
        "border": 0,
        "grad_color": ColorRange(Color(13421772), Color(6710886)),
        "grad_intensity": IntensityRange(100, 100),
        "name": PresetGradientKind.LONDON_MIST.value,
    }


def teal_blue() -> Dict[str, Any]:
    """Teal to Blue preset"""
    return {
        "style": GradientStyle.LINEAR,
        "step_count": 0,
        "offset": Offset(0, 0),
        "angle": 30,
        "border": 0,
        "grad_color": ColorRange(Color(5280650), Color(5866416)),
        "grad_intensity": IntensityRange(100, 100),
        "name": PresetGradientKind.TEAL_BLUE.value,
    }


def midnight() -> Dict[str, Any]:
    """Midnight preset"""
    return {
        "style": GradientStyle.LINEAR,
        "step_count": 0,
        "offset": Offset(0, 0),
        "angle": 0,
        "border": 0,
        "grad_color": ColorRange(Color(0), Color(2777241)),
        "grad_intensity": IntensityRange(100, 100),
        "name": PresetGradientKind.MIDNIGHT.value,
    }


def deep_ocean() -> Dict[str, Any]:
    """Deep Ocean preset"""
    return {
        "style": GradientStyle.RADIAL,
        "step_count": 0,
        "offset": Offset(50, 50),
        "angle": 0,
        "border": 0,
        "grad_color": ColorRange(Color(0), Color(7512015)),
        "grad_intensity": IntensityRange(100, 100),
        "name": PresetGradientKind.DEEP_OCEAN.value,
    }


def submarine() -> Dict[str, Any]:
    """Submarine preset"""
    return {
        "style": GradientStyle.LINEAR,
        "step_count": 0,
        "offset": Offset(0, 0),
        "angle": 0,
        "border": 0,
        "grad_color": ColorRange(Color(14543051), Color(11847644)),
        "grad_intensity": IntensityRange(100, 100),
        "name": PresetGradientKind.SUBMARINE.value,
    }


def green_grass() -> Dict[str, Any]:
    """Green Grass preset"""
    return {
        "style": GradientStyle.LINEAR,
        "step_count": 0,
        "offset": Offset(0, 0),
        "angle": 30,
        "border": 0,
        "grad_color": ColorRange(Color(16776960), Color(8508442)),
        "grad_intensity": IntensityRange(100, 100),
        "name": PresetGradientKind.GREEN_GRASS.value,
    }


def neon_light() -> Dict[str, Any]:
    """Neon Light preset"""
    return {
        "style": GradientStyle.ELLIPTICAL,
        "step_count": 0,
        "offset": Offset(50, 50),
        "angle": 0,
        "border": 15,
        "grad_color": ColorRange(Color(1209890), Color(16777215)),
        "grad_intensity": IntensityRange(100, 100),
        "name": PresetGradientKind.NEON_LIGHT.value,
    }


def sunshine() -> Dict[str, Any]:
    """Sunshine preset"""
    return {
        "style": GradientStyle.RADIAL,
        "step_count": 0,
        "offset": Offset(66, 33),
        "angle": 0,
        "border": 33,
        "grad_color": ColorRange(Color(16760576), Color(16776960)),
        "grad_intensity": IntensityRange(100, 100),
        "name": PresetGradientKind.SUNSHINE.value,
    }


def present() -> Dict[str, Any]:
    """Present preset"""
    return {
        "style": GradientStyle.SQUARE,
        "step_count": 0,
        "offset": Offset(70, 60),
        "angle": 45,
        "border": 72,
        "grad_color": ColorRange(Color(8468233), Color(16728064)),
        "grad_intensity": IntensityRange(100, 100),
        "name": PresetGradientKind.PRESENT.value,
    }


def mahogany() -> Dict[str, Any]:
    """Mahogany preset"""
    return {
        "style": GradientStyle.SQUARE,
        "step_count": 0,
        "offset": Offset(50, 50),
        "angle": 45,
        "border": 0,
        "grad_color": ColorRange(Color(0), Color(9250846)),
        "grad_intensity": IntensityRange(100, 100),
        "name": PresetGradientKind.MAHOGANY.value,
    }


def get_preset(kind: PresetGradientKind) -> Dict[str, Any]:
    """
    Gets preset

    Returns:
        PresetGradientKind: Preset Kind
    """
    if kind == PresetGradientKind.PASTEL_BOUQUET:
        return pastel_bouquet()
    if kind == PresetGradientKind.PASTEL_DREAM:
        return pastel_dream()
    if kind == PresetGradientKind.BLUE_TOUCH:
        return blue_touch()
    if kind == PresetGradientKind.BLANK_GRAY:
        return blank_gray()
    if kind == PresetGradientKind.SPOTTED_GRAY:
        return spotted_gray()
    if kind == PresetGradientKind.LONDON_MIST:
        return london_mist()
    if kind == PresetGradientKind.TEAL_BLUE:
        return teal_blue()
    if kind == PresetGradientKind.MIDNIGHT:
        return midnight()
    if kind == PresetGradientKind.DEEP_OCEAN:
        return deep_ocean()
    if kind == PresetGradientKind.SUBMARINE:
        return submarine()
    if kind == PresetGradientKind.GREEN_GRASS:
        return green_grass()
    if kind == PresetGradientKind.NEON_LIGHT:
        return neon_light()
    if kind == PresetGradientKind.SUNSHINE:
        return sunshine()
    if kind == PresetGradientKind.PRESENT:
        return present()
    if kind == PresetGradientKind.MAHOGANY:
        return mahogany()
