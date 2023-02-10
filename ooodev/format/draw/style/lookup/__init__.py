from __future__ import annotations
from typing import NamedTuple


class FamilyDefault(NamedTuple):
    """Draw Default family style names"""

    BACKGROUND = "background"
    BACKGROUND_OBJECTS = "backgroundobjects"
    NOTES = "notes"
    OUTLINE_1 = "outline1"
    OUTLINE_2 = "outline2"
    OUTLINE_3 = "outline3"
    OUTLINE_4 = "outline4"
    OUTLINE_5 = "outline5"
    OUTLINE_6 = "outline6"
    OUTLINE_7 = "outline7"
    OUTLINE_8 = "outline8"
    OUTLINE_9 = "outline9"
    SUBTITLE = "subtitle"
    TITLE = "title"

    @staticmethod
    def get_family_name() -> str:
        """Gets Family Name for lookup"""
        return "Default"


class FamilyCell(NamedTuple):
    """Draw Cell family default style names"""

    BLUE1 = "blue1"
    BLUE2 = "blue2"
    BLUE3 = "blue3"
    BW1 = "bw1"
    BW2 = "bw2"
    BW3 = "bw3"
    DEFAULT = "default"
    EARTH1 = "earth1"
    EARTH2 = "earth2"
    EARTH3 = "earth3"
    GRAY1 = "gray1"
    GRAY2 = "gray2"
    GRAY3 = "gray3"
    GREEN1 = "green1"
    GREEN2 = "green2"
    GREEN3 = "green3"
    LIGHTBLUE1 = "lightblue1"
    LIGHTBLUE2 = "lightblue2"
    LIGHTBLUE3 = "lightblue3"
    ORANGE1 = "orange1"
    ORANGE2 = "orange2"
    ORANGE3 = "orange3"
    SEETANG1 = "seetang1"
    SEETANG2 = "seetang2"
    SEETANG3 = "seetang3"
    SUN1 = "sun1"
    SUN2 = "sun2"
    SUN3 = "sun3"
    TURQUOISE1 = "turquoise1"
    TURQUOISE2 = "turquoise2"
    TURQUOISE3 = "turquoise3"
    YELLOW1 = "yellow1"
    YELLOW2 = "yellow2"
    YELLOW3 = "yellow3"

    @staticmethod
    def get_family_name() -> str:
        """Gets Family Name for lookup"""
        return "cell"


class FamilyGraphics(NamedTuple):
    """Draw Graphic family default style names"""

    A4 = "A4"
    A4 = "A4"
    DASHED_LINE = "Arrow Dashed"
    ARROW_LINE = "Arrow Line"
    FILLED = "Filled"
    FILLED_BLUE = "Filled Blue"
    FILLED_GREEN = "Filled Green"
    FILLED_RED = "Filled Red"
    FILLED_YELLOW = "Filled Yellow"
    GRAPHIC = "Graphic"
    HEADING_A0 = "Heading A0"
    HEADING_A4 = "Heading A4"
    LINES = "Lines"
    OBJECT_WITH_NO_FILL_AND_NO_LINE = "Object with no fill and no line"
    OUTLINED = "Outlined"
    OUTLINED_BLUE = "Outlined Blue"
    OUTLINED_GREEN = "Outlined Green"
    OUTLINED_RED = "Outlined Red"
    OUTLINED_YELLOW = "Outlined Yellow"
    SHAPES = "Shapes"
    TEXT = "Text"
    TEXT_A0 = "Text A0"
    TEXT_A4 = "Text A4"
    TITLE_A0 = "Title A0"
    TITLE_A4 = "Title A4"
    OBJECT_WITHOUT_FILL = "objectwithoutfill"
    DEFAULT_DRAWING_STYLE = "standard"

    @staticmethod
    def get_family_name() -> str:
        """Gets Family Name for lookup"""
        return "graphics"
