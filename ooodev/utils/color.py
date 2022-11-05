# coding: utf-8
"""
Various color conversions utilities.
"""
from __future__ import annotations
import math
import colorsys
from typing import Union, NamedTuple, overload, NewType
import numbers

from ..utils import gen_util as mGenUtil

# ref: https://gist.github.com/mathebox/e0805f72e7db3269ec22

MAX_COLOR = 255
"""Max Color Value"""
MIN_COLOR = 0
"""Min Color Value"""

Color = NewType("Color", int)
"""Color Type. Int RGB Value"""


class CommonColor:
    """
    Named Colors.

    See Also:
        - `Wikipedia Web colors <https://en.wikipedia.org/wiki/Web_colors>`_
        - `Hex Color Chart <https://www.quackit.com/css/color/charts/hex_color_chart.cfm>`_

    Example:

        .. code-block:: python

            def set_chart(sheet: XSpreadsheet, range_addr: CellRangeAddress) -> None:
                chart = Chart2.insert_chart(sheet, range_addr, "A22", 20, 11, Chart2.ChartLookup.Column.TEMPLATE_PERCENT.COLUMN_DEEP_3D)
                Chart2.set_background_colors(chart, CommonColor.LIGHT_GREEN, CommonColor.BROWN)
    """

    # https://en.wikipedia.org/wiki/Web_colors
    # https://www.quackit.com/css/color/charts/hex_color_chart.cfm
    # some hex values for commonly used colors
    # Pink colors
    DEEP_PINK = Color(0xFF1493)
    HOT_PINK = Color(0xFF69B4)
    LIGHT_PINK = Color(0xFFB6C1)
    MEDIUM_VIOLET_RED = Color(0xC71585)
    PALE_VIOLET_RED = Color(0xDB7093)
    PINK = Color(0xFFC0CB)

    # red colors
    CRIMSON = Color(0xDC143C)
    DARK_RED = Color(0x8B0000)
    DARK_SALMON = Color(0xE9967A)
    FIRE_BRICK = Color(0xB22222)
    INDIAN_RED = Color(0xCD5C5C)
    LIGHT_CORAL = Color(0xF08080)
    LIGHT_SALMON = Color(0xFFA07A)
    RED = Color(0xFF0000)
    SALMON = Color(0xFA8072)

    # Oranges
    CORAL = Color(0xFF7F50)
    DARK_ORANGE = Color(0xFF8C00)
    ORANGE = Color(0xFFA500)
    ORANGE_RED = Color(0xFF4500)
    TOMATO = Color(0xFF6347)

    # Yellow colors
    DARK_KHAKI = Color(0xBDB76B)
    GOLD = Color(0xFFD700)
    KHAKI = Color(0xF0E68C)
    LEMON_CHIFFON = Color(0xFFFACD)
    LIGHT_GOLDENROD_YELLOW = Color(0xFAFAD2)
    LIGHT_YELLOW = Color(0xFFFFE0)
    MOCCASIN = Color(0xFFE4B5)
    PALE_GOLDENROD = Color(0xEEE8AA)
    PAPAYA_WHIP = Color(0xFFEFD5)
    PEACH_PUFF = Color(0xFFDAB9)
    YELLOW = Color(0xFFFF00)

    # Purples
    BLUE_VIOLET = Color(0x8A2BE2)
    DARK_MAGENTA = Color(0x8B008B)
    DARK_ORCHID = Color(0x9932CC)
    DARK_SLATE_BLUE = Color(0x483D8B)
    DARK_VIOLET = Color(0x9400D3)
    FUCHSIA = Color(0xFF00FF)
    INDIGO = Color(0x4B0082)
    LAVENDER = Color(0xE6E6FA)
    MAGENTA = Color(0xFF00FF)
    MEDIUM_ORCHID = Color(0xBA55D3)
    MEDIUM_PURPLE = Color(0x9370DB)
    MEDIUM_SLATE_BLUE = Color(0x7B68EE)
    ORCHID = Color(0xDA70D6)
    PLUM = Color(0xDDA0DD)
    PURPLE = Color(0x800080)
    REBECCA_PURPLE = Color(0x663399)
    SLATE_BLUE = Color(0x6A5ACD)
    THISTLE = Color(0xD8BFD8)
    VIOLET = Color(0xEE82EE)

    # Greens
    CHARTREUSE = Color(0x7FFF00)
    DARK_CYAN = Color(0x008B8B)
    DARK_GREEN = Color(0x006400)
    DARK_OLIVE_GREEN = Color(0x556B2F)
    DARK_SEA_GREEN = Color(0x8FBC8F)
    FOREST_GREEN = Color(0x228B22)
    GREEN = Color(0x008000)
    GREEN_YELLOW = Color(0xADFF2F)
    LAWN_GREEN = Color(0x7CFC00)
    LIGHT_GREEN = Color(0x90EE90)
    LIGHT_SEA_GREEN = Color(0x20B2AA)
    LIME = Color(0x00FF00)
    LIME_GREEN = Color(0x32CD32)
    MEDIUM_AQUAMARINE = Color(0x66CDAA)
    MEDIUM_SEA_GREEN = Color(0x3CB371)
    MEDIUM_SPRING_GREEN = Color(0x00FA9A)
    OLIVE = Color(0x808000)
    OLIVE_DRAB = Color(0x6B8E23)
    PALE_GREEN = Color(0x98FB98)
    SEA_GREEN = Color(0x2E8B57)
    SPRING_GREEN = Color(0x00FF7F)
    TEAL = Color(0x008080)
    YELLOW_GREEN = Color(0x9ACD32)

    # Blues/Cyans
    AQUA = Color(0x00FFFF)
    AQUAMARINE = Color(0x7FFFD4)
    BLUE = Color(0x0000FF)
    CADET_BLUE = Color(0x5F9EA0)
    CORNFLOWER_BLUE = Color(0x6495ED)
    CYAN = Color(0x00FFFF)
    DARK_BLUE = Color(0x00008B)
    DARK_TURQUOISE = Color(0x00CED1)
    DEEP_SKY_BLUE = Color(0x00BFFF)
    DODGER_BLUE = Color(0x1E90FF)
    LIGHT_BLUE = Color(0xADD8E6)
    LIGHT_CYAN = Color(0xE0FFFF)
    LIGHT_SKY_BLUE = Color(0x87CEFA)
    LIGHT_STEEL_BLUE = Color(0xB0C4DE)
    MEDIUM_BLUE = Color(0x0000CD)
    MEDIUM_TURQUOISE = Color(0x48D1CC)
    MIDNIGHT_BLUE = Color(0x191970)
    NAVY = Color(0x000080)
    PALE_TURQUOISE = Color(0xAFEEEE)
    POWDER_BLUE = Color(0xB0E0E6)
    ROYAL_BLUE = Color(0x4169E1)
    SKY_BLUE = Color(0x87CEEB)
    STEEL_BLUE = Color(0x4682B4)
    TURQUOISE = Color(0x40E0D0)

    # Browns
    BISQUE = Color(0xFFE4C4)
    BLANCHED_ALMOND = Color(0xFFEBCD)
    BROWN = Color(0xA52A2A)
    BURLY_WOOD = Color(0xDEB887)
    CHOCOLATE = Color(0xD2691E)
    CORNSILK = Color(0xFFF8DC)
    DARK_GOLDENROD = Color(0xB8860B)
    GOLDENROD = Color(0xDAA520)
    MAROON = Color(0x800000)
    NAVAJO_WHITE = Color(0xFFDEAD)
    PERU = Color(0xCD853F)
    ROSY_BROWN = Color(0xBC8F8F)
    SADDLE_BROWN = Color(0x8B4513)
    SANDY_BROWN = Color(0xF4A460)
    SIENNA = Color(0xA0522D)
    TAN = Color(0xD2B48C)
    WHEAT = Color(0xF5DEB3)

    # Whites
    ALICE_BLUE = Color(0xF0F8FF)
    ANTIQUE_WHITE = Color(0xFAEBD7)
    AZURE = Color(0xF0FFFF)
    BEIGE = Color(0xF5F5DC)
    FLORAL_WHITE = Color(0xFFFAF0)
    GHOST_WHITE = Color(0xF8F8FF)
    HONEYDEW = Color(0xF0FFF0)
    IVORY = Color(0xFFFFF0)
    LAVENDER_BLUSH = Color(0xFFF0F5)
    LINEN = Color(0xFAF0E6)
    MINT_CREAM = Color(0xF5FFFA)
    MISTY_ROSE = Color(0xFFE4E1)
    OLD_LACE = Color(0xFDF5E6)
    SEASHELL = Color(0xFFF5EE)
    SNOW = Color(0xFFFAFA)
    WHITE = Color(0xFFFFFF)
    WHITE_SMOKE = Color(0xF5F5F5)

    # Greys
    BLACK = Color(0x000000)
    DARK_GRAY = Color(0xA9A9A9)
    DARK_GREY = Color(0xA9A9A9)
    DARK_SLATE_GRAY = Color(0x2F4F4F)
    DARK_SLATE_GREY = Color(0x2F4F4F)
    DIM_GRAY = Color(0x696969)
    DIM_GREY = Color(0x696969)
    GAINSBORO = Color(0xDCDCDC)
    GRAY = Color(0x808080)
    GREY = Color(0x808080)
    LIGHT_GRAY = Color(0xD3D3D3)
    LIGHT_GREY = Color(0xD3D3D3)
    LIGHT_SLATE_GRAY = Color(0x778899)
    LIGHT_SLATE_GREY = Color(0x778899)
    SILVER = Color(0xC0C0C0)
    SLATE_GRAY = Color(0x708090)
    SLATE_GREY = Color(0x708090)

    # other
    PALE_BLUE = Color(0xD6EBFF)

    @classmethod
    def from_str(cls, str_color: str) -> Color:
        """
        Convert string value to a Color value

        ``str_color`` can be a hex value, a integer value as string or any named value in ``CommonColor``

        ``str_color`` can contain single spaces or ``_`` or ``-``.
        ``MEDIUM_SEA_GREEN``, ``MEDIUM-SEA-GREEN``, ``MEDIUM SEA GREEN``, ``medium sea green``, ``mediumSeaGreen`` , ``MediumSeaGreen`` are
        all equivalent.

        Args:
            str_color (str): Value to convert. Case insensitive

        Raises:
            ValueError: If unable to convert.

        Returns:
            Color: Convert value as Color
        """
        if not str_color:
            raise ValueError("str_color contains no value to convert to Color")
        if str_color.isalnum():
            try:
                # try base 10
                i = int(str_color)
            except ValueError:
                # try hex
                try:
                    i = int(str_color, 16)
                except ValueError:
                    pass
                else:
                    return Color(abs(i))
            else:
                return Color(abs(i))
        c_str = mGenUtil.Util.to_single_space(str_color).replace("-", "_").replace(" ", "_")
        if "_" in c_str:
            c_str = c_str.upper()
        else:
            # handle pascalCase and CamelCase words
            c_str = mGenUtil.Util.to_snake_case_upper(c_str)
        try:
            return Color(int(getattr(cls, c_str)))
        except AttributeError:
            raise ValueError("str_color is not a valid color")


class RGB(NamedTuple):
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

    def to_color(self) -> Color:
        """
        Gets instance as rgb Color

        Returns:
            Color: red, green, blue encoded as Color.
        """
        return Color(self.to_int())

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
    def from_int(rgb_int: int) -> "RGB":
        """
        Gets a color instance from int that represents a rgb color.

        Args:
            rgb_int (int): int that contains rgb color data.

        Returns:
            RGB: Color information as RGB struct.
        """
        return int_to_rgb(rgb_int=rgb_int)

    @classmethod
    def from_color(cls, c: Color) -> "RGB":
        """
        Gets a color instance from input color that represents a rgb color.

        Args:
            c (Color): Color that contains rgb color data.

        Returns:
            RGB: Color information as RGB struct.
        """
        return cls.from_int(int(c))

    @staticmethod
    def from_hex(rgb_hex: str) -> "RGB":
        """
        Gets a color instance from int that represents a rgb color.

        Args:
            rgb_int (int): int that contains rgb color data.

        Returns:
            RGB: Color information as RGB struct.
        """
        return int_to_rgb(rgb_int=int(rgb_hex, 16))

    def get_luminance(self) -> float:
        """
        Gets luminance value for current color

        Returns:
            float: luminance value
        """
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
        Gets brightness from 0 (dark) to 255 (light)

        Returns:
            int: brightness level
        """
        # http://www.w3.org/TR/AERT#color-contrast
        return round(((self.red * 299) + (self.green * 587) + (self.blue * 114)) / 1000)

    def is_dark(self) -> bool:
        """
        Get is current color is dark.
        If color has a brightness less than ``128`` it is considered dark.

        Returns:
            bool: True if color is dark; Otherwise, False
        """
        return self.get_brightness() < 128

    def is_light(self) -> bool:
        """
        Get is current color is light.
        If color has a brightness Greater than ``128`` it is considered light.

        Returns:
            bool: True if color is light; Otherwise, False
        """
        return not self.is_dark()

    def __str__(self) -> str:
        return "rgb(" + round(self.red) + ", " + round(self.green) + ", " + round(self.blue) + ")"


class HSL(NamedTuple):
    hue: float
    saturation: float
    lightness: float

    def __str__(self) -> str:
        return "hls(" + f"{self.hue:.6f}" + ", " + f"{self.saturation:.6f}" + ", " + f"{self.lightness:.6f}" + ")"


class HSV(NamedTuple):
    hue: float
    saturation: float
    value: float

    def __str__(self) -> str:
        return "hlv(" + f"{self.hue:.6f}" + ", " + f"{self.saturation:.6f}" + ", " + f"{self.value:.6f}" + ")"


def clamp(value: float, min_value: float, max_value: float) -> float:
    """
    Constrains a value to a min and an max value

    Args:
        value (float): Value to constrain
        min_value (float): Min allowed value
        max_value (float): Max allowed value

    Returns:
        float: constrained value if value is outside of min_value or max_value; Otherwise, value.
    """
    return max(min_value, min(max_value, value))


def clamp01(value: float) -> float:
    """
    Gets a value that is constrained between 0.0 and 1.0

    Args:
        value (float): Value

    Returns:
        float: A value that is no less than 0.0 and no greater then 1.0
    """
    return clamp(value, 0.0, 1.0)


def hue_to_rgb(h: float) -> RGB:
    """
    Converts a hue to instance of red, green, blue

    Args:
        h (float): hue to convert

    Returns:
        rgb: instance containing red, green, blue
    """
    r = abs(h * 6.0 - 3.0) - 1.0
    g = 2.0 - abs(h * 6.0 - 2.0)
    b = 2.0 - abs(h * 6.0 - 4.0)
    return RGB(red=round(clamp01(r)), green=round(clamp01(g)), blue=round(clamp01(b)))


def hsl_to_rgb(c: HSL) -> RGB:
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
    return RGB(
        red=round(t[0] * MAX_COLOR),
        green=round(t[1] * MAX_COLOR),
        blue=round(t[2] * MAX_COLOR),
    )


def rgb_to_hsv(c: RGB) -> HSV:
    """
    Converts red, green, blue to hue, saturation, value

    Args:
        c (RGB): instance containing red, green, blue

    Returns:
        HSV: instance containing hue, saturation, value
    """
    r = float(c.red / MAX_COLOR)
    g = float(c.green / MAX_COLOR)
    b = float(c.blue / MAX_COLOR)
    t = colorsys.rgb_to_hsv(r=r, g=g, b=b)
    return HSV(hue=t[0], saturation=t[1], value=t[2])


def hsv_to_rgb(c: HSV) -> RGB:
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
    return RGB(
        red=round(t[0] * MAX_COLOR),
        green=round(t[1] * MAX_COLOR),
        blue=round(t[2] * MAX_COLOR),
    )


def rgb_to_hsl(c: RGB) -> HSL:
    """
    Converts red, green, blue to hue, saturation, value

    Args:
        c (RGB): instance containing red, green, blue

    Returns:
        HSL: instance containing hue, saturation, lightness
    """
    r = float(c.red / MAX_COLOR)
    g = float(c.green / MAX_COLOR)
    b = float(c.blue / MAX_COLOR)
    t = colorsys.rgb_to_hls(r=r, g=g, b=b)
    return HSL(hue=t[0], saturation=t[2], lightness=t[1])


def hsv_to_hsl(c: HSV) -> HSL:
    """
    Convert hue, saturation, value to hue, saturation, lightness

    Args:
        c (HSV): instance containing hue, saturation, value

    Returns:
        HSL: instance containing hue, saturation, lightness
    """
    h = c.hue
    s = c.saturation
    v = c.value
    l = 0.5 * v * (2 - s)
    s = v * s / (1 - math.fabs(2 * l - 1))
    return HSL(h, s, l)


def hsl_to_hsv(c: HSL) -> HSV:
    """
    Convert hue, saturation, lightness to hue, saturation, value

    Args:
        c (HSL): instance containing hue, saturation, lightness

    Returns:
        HSV: instance containing hue, saturation, value
    """
    h = c.hue
    s = c.saturation
    l = c.lightness
    v = (2 * l + s * (1 - math.fabs(2 * l - 1))) / 2
    s = 2 * (v - l) / v
    return HSV(h, s, v)


def rgb_to_hex(rgb: RGB) -> str:
    """
    Converts rgb colors to int

    Args:
        rgb (color): Tuple of int with values from 0 to 255

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


def rgb_to_int(rgb: RGB) -> int:
    """
    Converts rgb colors to int

    Args:
        rgb (color): Tuple of int with values from 0 to 255

    Returns:
        int: rgb as int
    """
    return int(rgb_to_hex(rgb), 16)


def int_to_rgb(rgb_int: int) -> RGB:
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
    return RGB(red, green, blue)


@overload
def lighten(rgb_color: int, percent: numbers.Number) -> RGB:
    ...


@overload
def lighten(rgb_color: RGB, percent: numbers.Number) -> RGB:
    ...


def lighten(rgb_color: Union[RGB, int], percent: numbers.Number) -> RGB:
    """
    Lightens an RGB instance

    Args:
        rgb_color (RGB | int): instance containing data
        percent (Number): Amount between 0 and 100 int lighten rgb by.

    Raises:
        ValueError: if percent is out of range

    Returns:
        RGB: RGB instance with lightened values applied.
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
    c2_hsl = HSL(c_hsl.hue, c_hsl.saturation, increase)
    c_rgb = hsl_to_rgb(c2_hsl)
    return c_rgb


@overload
def darken(rgb_color: int, percent: numbers.Number) -> RGB:
    ...


@overload
def darken(rgb_color: RGB, percent: numbers.Number) -> RGB:
    ...


def darken(rgb_color: Union[RGB, int], percent: numbers.Number) -> RGB:
    """
    Darkens an rgb instance

    Args:
        rgb_color (rgb | int): instance containing data
        percent (Number): Amount between 0 and 100 int darken rgb by.

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
    c2_hsl = HSL(c_hsl.hue, c_hsl.saturation, decrease)
    c_rgb = hsl_to_rgb(c2_hsl)
    return c_rgb
