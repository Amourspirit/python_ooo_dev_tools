from enum import IntEnum, unique
from ooodev.utils.kind import kind_helper


@unique
class PresentationLayoutKind(IntEnum):
    """
    ``com.sun.star.presentation.DrawPage`` service ``Layout`` property values.

    See Also:
        `API Presentation Drawpage Service <https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1presentation_1_1DrawPage.html>`_

    Example:
        .. code-block:: python

            Props.set(slide, Layout=PresentationLayoutKind.TITLE_ONLY.value)
    """

    TITLE_SUB = 0
    TITLE_BULLETS = 1
    TITLE_CHART = 2
    TITLE_2CONTENT = 3
    TITLE_CONTENT_CHART = 4
    TITLE_CONTENT_CLIP = 6
    TITLE_CHART_CONTENT = 7
    TITLE_TABLE = 8
    TITLE_CLIP_CONTENT = 9
    TITLE_CONTENT_OBJECT = 10
    TITLE_OBJECT = 11
    TITLE_CONTENT_2CONTENT = 12
    TITLE_OBJECT_CONTENT = 13
    TITLE_CONTENT_OVER_CONTENT = 14
    TITLE_2CONTENT_CONTENT = 15
    TITLE_2CONTENT_OVER_CONTENT = 16
    TITLE_CONTENT_OVER_OBJECT = 17
    TITLE_4OBJECT = 18
    TITLE_ONLY = 19
    BLANK = 20
    VTITLE_VTEXT_CHART = 27
    VTITLE_VTEXT = 28
    TITLE_VTEXT = 29
    TITLE_VTEXT_CLIP = 30
    CENTERED_TEXT = 32
    TITLE_4CONTENT = 33
    TITLE_6CONTENT = 34

    @staticmethod
    def from_str(s: str) -> "PresentationLayoutKind":
        """
        Gets an ``PresentationLayoutKind`` instance from string.

        Args:
            s (str): String that represents the name of an enum Attribute or enum value as string.
                ``s`` is case insensitive and can be ``CamelCase``, ``pascal_case`` , ``snake_case``,
                ``hyphen-case``, ``normal case``.

        Raises:
            ValueError: If input string is empty.
            AttributeError: If unable to get ``PresentationLayoutKind`` instance.

        Returns:
            PresentationLayoutKind: Enum instance.
        """
        return kind_helper.enum_from_string(s, PresentationLayoutKind)
