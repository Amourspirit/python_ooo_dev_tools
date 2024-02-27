from enum import Enum
from ooodev.utils.kind import kind_helper

# these are helper lookups
# Example Usage:
#   Draw.set_style(shape=conn_shape, graphic_styles=g_styles, style_name=GraphicStyleKind.ARROW_LINE)
#
# see the Grouper Sample.
# value can be obtaind in the follow mannor.
# g_styles = Info.get_style_container(doc=doc, family_style_name="graphics")
# Info.show_container_names("Graphic styles", g_styles)


class GraphicStyleKind(str, Enum):
    """
    Graphic Styles
    """

    STANDARD = "standard"
    """The style Default (standard) is used for newly inserted filled rectangles, filled ellipses, lines, connectors, text boxes, and 3D objects."""
    OBJECT_WITHOUT_FILL = "objectwithoutfill"
    OBJECT_WITH_NO_FILL_AND_NO_LINE = "Object with no fill and no line"
    TEXT = "Text"
    A4 = "A4"
    TITLE_A4 = "Title A4"
    HEADING_A4 = "Heading A4"
    TEXT_A4 = "Text A4"
    TITLE_A0 = "Title A0"
    HEADING_A0 = "Heading A0"
    TEXT_A0 = "Text A0"
    GRAPHIC = "Graphic"
    SHAPES = "Shapes"
    FILLED = "Filled"
    FILLED_BLUE = "Filled Blue"
    FILLED_GREEN = "Filled Green"
    FILLED_RED = "Filled Red"
    FILLED_YELLOW = "Filled Yellow"
    OUTLINED = "Outlined"
    OUTLINED_BLUE = "Outlined Blue"
    OUTLINED_GREEN = "Outlined Green"
    OUTLINED_RED = "Outlined Red"
    OUTLINED_YELLOW = "Outlined Yellow"
    LINES = "Lines"
    ARROW_LINE = "Arrow Line"
    ARROW_DASHED = "Arrow Dashed"

    def __str__(self) -> str:
        return self.value

    @staticmethod
    def from_str(s: str) -> "GraphicStyleKind":
        """
        Gets an ``GraphicStyleKind`` instance from string.

        Args:
            s (str): String that represents the name of an enum Name.
                ``s`` is case insensitive and can be ``CamelCase``, ``pascal_case`` , ``snake_case``,
                ``hyphen-case``, ``normal case``.

        Raises:
            ValueError: If input string is empty.
            AttributeError: If unable to get ``GraphicStyleKind`` instance.

        Returns:
            GraphicStyleKind: Enum instance.
        """
        return kind_helper.enum_from_string(s, GraphicStyleKind)
