from enum import Enum

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
    DIAMOND_UNFILLED = "Daimond unfilled"
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
    SQURE_UNFILLED = "Squre unfilled"
    TRIANGLE = "Triangle"
    TRIANGLE_UNFILLED = "Triangle unfilled"

    def __str__(self) -> str:
        return self.value
