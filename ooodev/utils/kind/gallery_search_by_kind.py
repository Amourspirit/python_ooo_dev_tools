from enum import IntEnum
from ooodev.utils.kind import kind_helper


class SearchByKind(IntEnum):
    """
    Gallery Search by Kind.

    Used to determine search criteria.

    See Also:
        :py:meth:`.Gallery.find_gallery_item`
    """

    FILE_NAME = 1
    """Match File Name"""
    TITLE = 2
    """Match Title"""

    @staticmethod
    def from_str(s: str) -> "SearchByKind":
        """
        Gets an ``SearchByKind`` instance from string.

        Args:
            s (str): String that represents the name of an enum Name.
                ``s`` is case insensitive and can be ``CamelCase``, ``pascal_case`` , ``snake_case``,
                ``hyphen-case``, ``normal case``.

        Raises:
            ValueError: If input string is empty.
            AttributeError: If unable to get ``SearchByKind`` instance.

        Returns:
            SearchByKind: Enum instance.
        """
        return kind_helper.enum_from_string(s, SearchByKind)
