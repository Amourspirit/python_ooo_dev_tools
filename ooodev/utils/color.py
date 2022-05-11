# coding: utf-8
"""
Various color conversions utilities.
"""
import math
import colorsys
from typing import Union, NamedTuple, overload

# ref: https://gist.github.com/mathebox/e0805f72e7db3269ec22

MAX_COLOR = 255
MIN_COLOR = 0


class rgb(NamedTuple):
    red: int
    """Red color as int"""
    green: int
    """Green color as int"""
    blue: int
    """Blue color as int"""

    def to_int(self) -> int:
        """
        Gets instance as rgb int

        Returns:
            int: red, green, blue encoded as int.
        """
        return rgb_to_int(self)

    def to_hex(self) -> str:
        """
        Gets instance as hex string in format of ``ff3322``

        Returns:
            str: red, green, blue encoded as hex string.
        """
        return rgb_to_hex(self)

    def isvalid(self) -> bool:
        """
        Gets if the value of red, green and blue are valid.

        Returns:
            bool: ``True`` if valid; Otherwise, ``False``
        """
        result = True
        for i in self:
            if i < MIN_COLOR or i > MAX_COLOR:
                result = False
                break
        return result

    @staticmethod
    def from_int(rgb_int: int) -> "rgb":
        """
        Gets a color instance from int that represents a rgb color.

        Args:
            rgb_int (int): int that contains rgb color data.

        Returns:
            color: color struct.
        """
        return int_to_rgb(rgb_int=rgb_int)

    @staticmethod
    def from_hex(rgb_hex: str) -> "rgb":
        """
        Gets a color instance from int that represents a rgb color.

        Args:
            rgb_int (int): int that contains rgb color data.

        Returns:
            color: color struct.
        """
        return int_to_rgb(rgb_int=int(rgb_hex, 16))

    def get_luminance(self) -> float:
        # http://www.w3.org/TR/2008/REC-WCAG20-20081211/#relativeluminancedef

        RsRGB = self.red / 255
        GsRGB = self.green / 255
        BsRGB = self.blue / 255

        if RsRGB <= 0.03928:
            R = RsRGB / 12.92
        else:
            R = math.pow(((RsRGB + 0.055) / 1.055), 2.4)
        if GsRGB <= 0.03928:
            G = GsRGB / 12.92
        else:
            G = math.pow(((GsRGB + 0.055) / 1.055), 2.4)
        if BsRGB <= 0.03928:
            B = BsRGB / 12.92
        else:
            B = math.pow(((BsRGB + 0.055) / 1.055), 2.4)
        return (0.2126 * R) + (0.7152 * G) + (0.0722 * B)

    def get_brightness(self) -> int:
        """
        Gets brightnes from 0 (dark) to 255 (light)

        Returns:
            int: brightness level
        """
        # http://www.w3.org/TR/AERT#color-contrast
        return round(((self.red * 299) + (self.green * 587) + (self.blue * 114)) / 1000)

    def is_dark(self) -> bool:
        return self.get_brightness() < 128

    def is_light(self) -> bool:
        return not self.is_dark()

    def __str__(self) -> str:
        return (
            "rgb("
            + round(self.red)
            + ", "
            + round(self.green)
            + ", "
            + round(self.blue)
            + ")"
        )


class hsl(NamedTuple):
    hue: float
    saturation: float
    lightness: float

    def __str__(self) -> str:
        return (
            "hls("
            + f"{self.hue:.6f}"
            + ", "
            + f"{self.saturation:.6f}"
            + ", "
            + f"{self.lightness:.6f}"
            + ")"
        )


class hsv(NamedTuple):
    hue: float
    saturation: float
    value: float

    def __str__(self) -> str:
        return (
            "hlv("
            + f"{self.hue:.6f}"
            + ", "
            + f"{self.saturation:.6f}"
            + ", "
            + f"{self.value:.6f}"
            + ")"
        )


def clamp(value: float, min_value: float, max_value: float) -> float:
    return max(min_value, min(max_value, value))


def clamp01(value: float) -> float:
    return clamp(value, 0.0, 1.0)


def hue_to_rgb(h: float) -> rgb:
    """
    Conerts a hue to instance of red, gree, blue

    Args:
        h (float): hue to convert

    Returns:
        rgb: instance containing red, green, blue
    """
    r = abs(h * 6.0 - 3.0) - 1.0
    g = 2.0 - abs(h * 6.0 - 2.0)
    b = 2.0 - abs(h * 6.0 - 4.0)
    return rgb(red=round(clamp01(r)), green=round(clamp01(g)), blue=round(clamp01(b)))


def hsl_to_rgb(c: hsl) -> rgb:
    """
    Converts hue, saturation, lightness to red, green, blue

    Args:
        c (hsv): instance containing hue, saturation, lightness

    Returns:
        rgb: instance containing red, green, blue
    """
    h = c.hue
    l = c.lightness
    s = c.saturation

    t = colorsys.hls_to_rgb(h=h, l=l, s=s)
    return rgb(
        red=round(t[0] * MAX_COLOR),
        green=round(t[1] * MAX_COLOR),
        blue=round(t[2] * MAX_COLOR),
    )


def rgb_to_hsv(c: rgb) -> hsv:
    """
    Converts red, green, blue to hue, saturation, value

    Args:
        c (rgb): instance containing red, green, blue

    Returns:
        hsv: instance containing hue, saturation, value
    """
    r = float(c.red / MAX_COLOR)
    g = float(c.green / MAX_COLOR)
    b = float(c.blue / MAX_COLOR)
    t = colorsys.rgb_to_hsv(r=r, g=g, b=b)
    return hsv(hue=t[0], saturation=t[1], value=t[2])


def hsv_to_rgb(c: hsv) -> rgb:
    """
    Converts hue, saturation, value to red, green, blue

    Args:
        c (hsv): instance containing hue, saturation, value

    Returns:
        rgb: instance containing red, green, blue
    """
    h = c.hue
    s = c.saturation
    v = c.value
    t = colorsys.hsv_to_rgb(h=h, s=s, v=v)
    return rgb(
        red=round(t[0] * MAX_COLOR),
        green=round(t[1] * MAX_COLOR),
        blue=round(t[2] * MAX_COLOR),
    )


def rgb_to_hsl(c: rgb) -> hsl:
    """
    Converts red, green, blue to hue, saturation, value

    Args:
        c (rgb): instance containing red, green, blue

    Returns:
        hsl: instance containing hue, saturation, lightness
    """
    r = float(c.red / MAX_COLOR)
    g = float(c.green / MAX_COLOR)
    b = float(c.blue / MAX_COLOR)
    t = colorsys.rgb_to_hls(r=r, g=g, b=b)
    return hsl(hue=t[0], saturation=t[2], lightness=t[1])


def hsv_to_hsl(c: hsv) -> hsl:
    """
    Convert hue, saturation, value to hue, saturation, lightness

    Args:
        c (hsv): instance containing hue, saturation, value

    Returns:
        hsl: instance containing hue, saturation, lightness
    """
    h = c.hue
    s = c.saturation
    v = c.value
    l = 0.5 * v * (2 - s)
    s = v * s / (1 - math.fabs(2 * l - 1))
    return hsl(h, s, l)


def hsl_to_hsv(c: hsl) -> hsv:
    """
    Convert hue, saturation, lightness to hue, saturation, value

    Args:
        c (hsl): instance containing hue, saturation, lightness

    Returns:
        hsv: instance containing hue, saturation, value
    """
    h = c.hue
    s = c.saturation
    l = c.lightness
    v = (2 * l + s * (1 - math.fabs(2 * l - 1))) / 2
    s = 2 * (v - l) / v
    return hsv(h, s, v)


def rgb_to_hex(rgb: rgb) -> str:
    """
    Converts rgb colors to int

    Args:
        rgb (color): Tuple of int with values from 0 to MAX_COLOR

    Returns:
        str: rgb as hex string
    """
    if len(rgb) != 3:
        raise ValueError("rgb must be a tuple of 3 integers")
    for i in rgb:
        if i < 0:
            raise ValueError("rgb contains a negative value")
        if i > MAX_COLOR:
            raise ValueError("rgb contains a value that is greater than MAX_COLOR")
    return "%02x%02x%02x" % rgb


def rgb_to_int(rgb: rgb) -> int:
    """
    Converts rgb colors to int

    Args:
        rgb (color): Tuple of int with values from 0 to MAX_COLOR

    Returns:
        int: rgb as int
    """
    return int(rgb_to_hex(rgb), 16)


def int_to_rgb(rgb_int: int) -> rgb:
    """
    Converts an integer that represents a rgb color into rgb object.

    Args:
        rgb_int (int): int that represents rgb color

    Returns:
        rgb: rgb with red, green and blue properties.
    """
    blue = rgb_int & MAX_COLOR
    green = (rgb_int >> 8) & MAX_COLOR
    red = (rgb_int >> 16) & MAX_COLOR
    return rgb(red, green, blue)


@overload
def lighten(rgb_color: int, percent: float) -> rgb:
    ...


@overload
def lighten(rgb_color: rgb, percent: float) -> rgb:
    ...


@overload
def lighten(rgb_color: int, percent: int) -> rgb:
    ...


@overload
def lighten(rgb_color: rgb, percent: int) -> rgb:
    ...


def lighten(rgb_color: Union[rgb, int], percent: Union[float, int]) -> rgb:
    """
    Lightenes an rgb instance

    Args:
        rgb_color (Union[rgb, int]): instanct containing data
        percent (Union[float, int]): Amount between 0 and 100 int lighten rgb by.

    Raises:
        ValueError: if percent is out of range

    Returns:
        rgb: rgb instance with lightened values applied.
    """
    if percent < 0 or percent > 100:
        raise ValueError("percent is expected to be between 0 and 100")
    # https://mdigi.tools/lighten-color
    # https://pastebin.com/KBAbAPh0
    if isinstance(rgb_color, int):
        _rgb = int_to_rgb(rgb_color)
    else:
        _rgb = rgb_color
    c_hsl = rgb_to_hsl(_rgb)
    amt = (percent * ((1 - c_hsl.lightness) * 100)) / 100
    l = c_hsl.lightness
    l += amt / 100
    increase = clamp(l, 0, 1)
    c2_hsl = hsl(c_hsl.hue, c_hsl.saturation, increase)
    c_rgb = hsl_to_rgb(c2_hsl)
    return c_rgb


@overload
def darken(rgb_color: int, percent: float) -> rgb:
    ...


@overload
def darken(rgb_color: rgb, percent: float) -> rgb:
    ...


@overload
def darken(rgb_color: int, percent: int) -> rgb:
    ...


@overload
def darken(rgb_color: rgb, percent: int) -> rgb:
    ...


def darken(rgb_color: Union[rgb, int], percent: Union[float, int]) -> rgb:
    """
    Darkens an rgb instance

    Args:
        rgb_color (Union[rgb, int]): instanct containing data
        percent (Union[float, int]): Amount between 0 and 100 int darken rgb by.

    Raises:
        ValueError: if percent is out of range

    Returns:
        rgb: rgb instance with darkened values applied.
    """
    if percent < 0 or percent > 100:
        raise ValueError("percent is expected to be between 0 and 100")
    # https://mdigi.tools/lighten-color
    # https://pastebin.com/KBAbAPh0
    if isinstance(rgb_color, int):
        _rgb = int_to_rgb(rgb_color)
    else:
        _rgb = rgb_color
    c_hsl = rgb_to_hsl(_rgb)
    amt = (percent * (c_hsl.lightness * 100)) / 100
    l = c_hsl.lightness
    l -= amt / 100
    decrease = clamp(l, 0, 1)
    c2_hsl = hsl(c_hsl.hue, c_hsl.saturation, decrease)
    c_rgb = hsl_to_rgb(c2_hsl)
    return c_rgb
