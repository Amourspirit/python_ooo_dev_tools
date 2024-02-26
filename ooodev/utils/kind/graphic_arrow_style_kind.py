from enum import Enum
from ooodev.utils.kind import kind_helper

# see Grouper Example


class GraphicArrowStyleKind(str, Enum):
    """
    Graphic Arrow Styles Lookup Values

    The values here are sourced from LibreOffice 7.4
    Draw, Graphic Styles: Arrow Line dialog box.

    Example:

        .. code-block:: python

            from ooodev.office.draw import Draw GraphicStyleKind
            from ooodev.utils.props import Props
            from ooodev.utils.kind.graphic_arrow_style_kind import GraphicArrowStyleKind

            Draw.set_style(shape=conn_shape, graphic_styles=g_styles, style_name=GraphicStyleKind.ARROW_LINE)

            # note that GraphicArrowStyleKind.ARROW_SHORT is wrapped in str() to get the required string value.
            Props.set(
                conn_shape,
                LineWidth=50,
                FillColor=CommonColor.DARK_BLUE,
                LineStartName=str(GraphicArrowStyleKind.ARROW_SHORT),
                LineStartCenter=False,
                LineEndName=str(GraphicArrowStyleKind.DIAMOND),
            )
    """

    NONE = ""
    ARROW = "Arrow"
    ARROW_LARGE = "Arrow large"
    ARROW_SHORT = "Arrow short"
    CF_MANY = "CF Many"
    CF_MANY_ONE = "CF Many One"
    CF_ONE = "CF One"
    CF_ONLY_ONE = "CF Only One"
    CF_ZERO_MANY = "CF Zero Many"
    CF_ZERO_ONE = "CF Zero One"
    CIRCLE = "Circle"
    CIRCLE_UNFILLED = "Circle unfilled"
    CONCAVE = "Concave"
    CONCAVE_SHORT = "Concave short"
    DIAMOND = "Diamond"
    DIAMOND_UNFILLED = "Diamond unfilled"
    DIMENSION_LINE = "Dimension Line"
    DIMENSION_LINE_ARROW = "Dimension Line Arrow"
    DIMENSION_LINES = "Dimension Lines"
    DOUBLE_ARROW = "Double Arrow"
    HALF_ARROW_LEFT = "Half Arrow left"
    HALF_ARROW_RIGHT = "Half Arrow right"
    HALF_CIRCLE = "Half Circle"
    HALF_CIRCLE_UNFILLED = "Half Circle unfilled"
    LINE = "Line"
    LINE_SHORT = "Line short"
    REVERSE_ARROW = "Reverse Arrow"
    SQUARE = "Square"
    SQUARE_45 = "Square 45"
    SQUARE_45_UNFILLED = "Square 45 unfilled"
    SQURE_UNFILLED = "Square unfilled"
    """Same as SQUARE_UNFILLED"""
    SQUARE_UNFILLED = "Square unfilled"
    TRIANGLE = "Triangle"
    TRIANGLE_UNFILLED = "Triangle unfilled"

    def __str__(self) -> str:
        return self.value

    @staticmethod
    def from_str(s: str) -> "GraphicArrowStyleKind":
        """
        Gets an ``GraphicArrowStyleKind`` instance from string.

        Args:
            s (str): String that represents the name of an enum Name.
                ``s`` is case insensitive and can be ``CamelCase``, ``pascal_case`` , ``snake_case``,
                ``hyphen-case``, ``normal case``.

        Raises:
            ValueError: If input string is empty.
            AttributeError: If unable to get ``GraphicArrowStyleKind`` instance.

        Returns:
            GraphicArrowStyleKind: Enum instance.
        """
        return kind_helper.enum_from_string(s, GraphicArrowStyleKind)
