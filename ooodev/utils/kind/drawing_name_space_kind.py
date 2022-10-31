from enum import Enum


class DrawingNameSpaceKind(str, Enum):
    """
    Draw Namespace kind.

    See Also:
        :py:meth:`.Draw.find_shape_by_type`
    """

    BULLETS_TEXT = "com.sun.star.presentation.OutlinerShape"
    SHAPE_TYPE_FOOTER = "com.sun.star.presentation.FooterShape"
    SHAPE_TYPE_NOTES = "com.sun.star.presentation.NotesShape"
    SHAPE_TYPE_PAGE = "com.sun.star.presentation.PageShape"
    SUBTITLE_TEXT = "com.sun.star.presentation.SubtitleShape"
    TITLE_TEXT = "com.sun.star.presentation.TitleTextShape"
    OLE2_SHAPE = "com.sun.star.drawing.OLE2Shape"

    def __str__(self) -> str:
        return self.value
