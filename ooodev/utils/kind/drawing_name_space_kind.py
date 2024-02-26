from enum import Enum
from ooodev.utils.kind import kind_helper


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

    @staticmethod
    def from_str(s: str) -> "DrawingNameSpaceKind":
        """
        Gets an ``DrawingNameSpaceKind`` instance from string.

        Args:
            s (str): String that represents the name of an enum Name.
                ``s`` is case insensitive and can be ``CamelCase``, ``pascal_case`` , ``snake_case``,
                ``hyphen-case``, ``normal case``.

        Raises:
            ValueError: If input string is empty.
            AttributeError: If unable to get ``DrawingNameSpaceKind`` instance.

        Returns:
            DrawingNameSpaceKind: Enum instance.
        """
        return kind_helper.enum_from_string(s, DrawingNameSpaceKind)
