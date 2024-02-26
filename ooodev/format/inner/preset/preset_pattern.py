from __future__ import annotations
from enum import Enum
from typing import ByteString
import uno
from com.sun.star.awt import XBitmap

from ooodev.utils.images_lo import BitmapArgs
from ooodev.utils.images_lo import ImagesLo


# region Enum
class PresetPatternKind(Enum):
    """Pattern Kind"""

    DASHED_DOWNWARD_DIAGONAL = "Dashed Downward Diagonal"
    """Dashed Downward Diagonal"""
    DASHED_DOTTED_UPWARD_DIAGONAL = "Dashed Dotted Upward Diagonal"
    """Dashed Dotted Upward Diagonal"""
    DASHED_HORIZONTAL = "Dashed Horizontal"
    """Dashed Horizontal"""
    DIAGONAL_BRICK = "Diagonal Brick"
    """Diagonal Brick"""
    DIVOT = "Divot"
    """Divot"""
    DOTTED_GRID = "Dotted Grid"
    """Dotted Grid"""
    HORIZONTAL_BRICK = "Horizontal Brick"
    """Horizontal Brick"""
    LARGE_CONFETTI = "Large Confetti"
    """Large Confetti"""
    PERCENT_10 = "10 Percent"
    """10 Percent"""
    PERCENT_20 = "20 Percent"
    """20 Percent"""
    PERCENT_5 = "5 Percent"
    """5 Percent"""
    SHINGLE = "Shingle"
    """Shingle"""
    SPHERE = "Sphere"
    """Sphere"""
    WEAVE = "Weave"
    """Weave"""
    ZIG_ZAG = "Zig Zag"
    """Zig Zag"""

    def __str__(self) -> str:
        return self.value


# endregion Enum


# region Base64 Pattern Images
def _get_b64_horizontal_brick():
    """Horizontal Brick"""
    return b"iVBORw0KGgoAAAANSUhEUgAAAAgAAAAIAQMAAAD+wSzIAAAABlBMVEX/e1m+SApuEXz8AAAAFElEQVR4nGNgYmBi+M+gAIb/GZgAENwCZQ+9fl8AAAAASUVORK5CYII="


def _get_b64_divot():
    """Divot"""
    return b"iVBORw0KGgoAAAANSUhEUgAAAAgAAAAIAQMAAAD+wSzIAAAABlBMVEV3vGX///9qMUCAAAAAFklEQVR4nGNwYFBgcGBgYGBiYAFiBgAIWACpMQYK2wAAAABJRU5ErkJggg=="


def _get_b64_large_confetti():
    """Large Confetti"""
    return b"iVBORw0KGgoAAAANSUhEUgAAAAgAAAAIAQMAAAD+wSzIAAAABlBMVEX//efgQPvJW0tfAAAAGElEQVR4nGOQYGhk2MhgxsDGkMCQzCANABTUAmVnPXTMAAAAAElFTkSuQmCC"


def _get_b64_dashed_horzontal():
    """Dashed Horizontal"""
    return b"iVBORw0KGgoAAAANSUhEUgAAAAgAAAAIAQMAAAD+wSzIAAAABlBMVEXo8qFeuR4xeXZVAAAAEklEQVR4nGNgAAI5Bgh4yMAAAAP9AQDzy0uqAAAAAElFTkSuQmCC"


def _get_b64_sphere():
    """Sphere"""
    return b"iVBORw0KGgoAAAANSUhEUgAAAAgAAAAIAQMAAAD+wSzIAAAABlBMVEXd3d1mZmZ1sNNMAAAAFUlEQVR4nGOYyyADhEoMNxkOAqESABkuA3WYLtF2AAAAAElFTkSuQmCC"


def _get_b64_weave():
    """Weave"""
    return b"iVBORw0KGgoAAAANSUhEUgAAAAgAAAAIAQMAAAD+wSzIAAAABlBMVEX/vwD/9c54u0NkAAAAFklEQVR4nGMIZOhgEGVQYggA0q4MSgATmQJQa7tR1gAAAABJRU5ErkJggg=="


def _get_b64_zig_zag():
    """Zig Zag"""
    return b"iVBORw0KGgoAAAANSUhEUgAAAAgAAAAIAQMAAAD+wSzIAAAABlBMVEXe5u8qYJl9LyTwAAAAEklEQVR4nGPwYGhhYGYwYIDSABGSAf+f4Tq6AAAAAElFTkSuQmCC"


def _get_b64_shingle():
    """Shingle"""
    return b"iVBORw0KGgoAAAANSUhEUgAAAAgAAAAIAQMAAAD+wSzIAAAABlBMVEX/QAD/qpVzRU0PAAAAFklEQVR4nGPgYGABQh4GIQZFhgMMBgAE5wFA1/K39wAAAABJRU5ErkJggg=="


def _get_b64_digonal_brick():
    """Diagonal Brick"""
    return b"iVBORw0KGgoAAAANSUhEUgAAAAgAAAAIAQMAAAD+wSzIAAAABlBMVEVmZmbd3d06UMziAAAAGElEQVR4nGNwYGhgYGRgZmhh8GAwYFAAABAqAeH+E7mWAAAAAElFTkSuQmCC"


def _get_b64_dashed_dotted_upward_diagonal():
    """Dashed Dotted Upward Diagonal"""
    return b"iVBORw0KGgoAAAANSUhEUgAAAAgAAAAIAQMAAAD+wSzIAAAABlBMVEX///8AAABVwtN+AAAAFUlEQVR4nGNgYGBiAAEBBgUGBwYGAAH6AHOiMCeUAAAAAElFTkSuQmCC"


def _get_b64_dotted_grid():
    """Dotted Grid"""
    return b"iVBORw0KGgoAAAANSUhEUgAAAAgAAAAIAQMAAAD+wSzIAAAABlBMVEX///8AAABVwtN+AAAAE0lEQVQI12NgYGBiYGBYxQChmQAGMACxp9kvAgAAAABJRU5ErkJggg=="


def _get_b64_dashed_downward_diogonal():
    """Dashed Downward Diagonal"""
    return b"iVBORw0KGgoAAAANSUhEUgAAAAgAAAAIAQMAAAD+wSzIAAAABlBMVEX///8AAABVwtN+AAAAFElEQVQI12NgYOhgcGFQYhBkAAMAC40BAEUuU08AAAAASUVORK5CYII="


def _get_b64_20():
    """20 Percent"""
    return b"iVBORw0KGgoAAAANSUhEUgAAAAgAAAAIAQMAAAD+wSzIAAAABlBMVEX///8AAABVwtN+AAAAEklEQVQI12NgYHBhYGAQZIDSAAWCAKvnQf5VAAAAAElFTkSuQmCC"


def _get_b64_10():
    """10 Percent"""
    return b"iVBORw0KGgoAAAANSUhEUgAAAAgAAAAIAQMAAAD+wSzIAAAABlBMVEX///8AAABVwtN+AAAAEklEQVR4nGNgYGBiYGBQYIDSAAF0AEUmPTVuAAAAAElFTkSuQmCC"


def _get_b64_5():
    """5 Percent"""
    return b"iVBORw0KGgoAAAANSUhEUgAAAAgAAAAIAQMAAAD+wSzIAAAABlBMVEX///8AAABVwtN+AAAAEklEQVR4nGNgAAIFBghgYmAAAAF2ACN1w4S1AAAAAElFTkSuQmCC"


def _get_b64(preset: PresetPatternKind) -> ByteString:
    if preset == PresetPatternKind.DASHED_DOWNWARD_DIAGONAL:
        return _get_b64_dashed_downward_diogonal()
    if preset == PresetPatternKind.DASHED_DOTTED_UPWARD_DIAGONAL:
        return _get_b64_dashed_dotted_upward_diagonal()
    if preset == PresetPatternKind.DASHED_HORIZONTAL:
        return _get_b64_dashed_horzontal()
    if preset == PresetPatternKind.DIAGONAL_BRICK:
        return _get_b64_digonal_brick()
    if preset == PresetPatternKind.DIVOT:
        return _get_b64_divot()
    if preset == PresetPatternKind.DOTTED_GRID:
        return _get_b64_dotted_grid()
    if preset == PresetPatternKind.HORIZONTAL_BRICK:
        return _get_b64_horizontal_brick()
    if preset == PresetPatternKind.LARGE_CONFETTI:
        return _get_b64_large_confetti()
    if preset == PresetPatternKind.PERCENT_10:
        return _get_b64_10()
    if preset == PresetPatternKind.PERCENT_20:
        return _get_b64_20()
    if preset == PresetPatternKind.PERCENT_5:
        return _get_b64_5()
    if preset == PresetPatternKind.SHINGLE:
        return _get_b64_shingle()
    if preset == PresetPatternKind.SPHERE:
        return _get_b64_sphere()
    if preset == PresetPatternKind.WEAVE:
        return _get_b64_weave()
    return _get_b64_zig_zag()


def get_prest_bitmap(preset: PresetPatternKind) -> XBitmap:
    """
    Gets preset Image

    Args:
        preset (ImageKind): Present image to return

    Returns:
        XBitmap: Preset Image
    """
    b64 = _get_b64(preset)
    bargs = BitmapArgs(name=str(preset), auto_name=False, auto_update=False)
    return ImagesLo.get_bitmap_from_b64(b64, bargs)


# endregion Base64 Pattern Images
