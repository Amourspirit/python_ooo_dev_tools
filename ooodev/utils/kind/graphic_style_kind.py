from enum import Enum


class GraphicStyleKind(str, Enum):
    """
    Graphic Styles

    See Also:
        `Developers Guide Styles <https://wiki.openoffice.org/wiki/Documentation/DevGuide/Drawings/Overall_Document_Features>`_
    """

    DEFAULT = "standard"
    """The style Default (standard) is used for newly inserted filled rectangles, filled ellipses, lines, connectors, text boxes, and 3D objects."""
    DIMENSION_LINE = "measure"
    """Used for newly inserted dimension lines."""
    FIRST_LINE_INDENT = "textbodyindent"
    HEADING = "headline"
    HEADING1 = "headline1"
    HEADING2 = "headline2"
    OBJECT_WITH_ARROW = "objectwitharrow"
    OBJECT_WITH_SHADOW = "objectwithshadow"
    OBJECT_WITHOUT_FILL = "objectwithoutfill"
    TEXT = "text"
    """Newly inserted text boxes do not use this style. They use Default and remove the fill settings for Default."""
    TEXT_BODY = "textbody"
    TEXT_BODY_JUSTIFIED = "textbodyjustfied"
    TITLE = "title"
    TITLE1 = "title1"
    TITLE2 = "title2"

    def __str__(self) -> str:
        return self.value
