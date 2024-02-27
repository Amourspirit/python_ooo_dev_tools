from enum import IntEnum
from ooodev.utils.kind import kind_helper


class SearchMatchKind(IntEnum):
    """Search Match Kind"""

    FULL = 1
    """Full search match. Search criteria must match Exact"""
    FULL_IGNORE_CASE = 2
    """Full search match. Case is ignored. Search criteria must match Exact"""
    PARTIAL = 3
    """Partial match. Search criteria must be in the search result"""
    PARTIAL_IGNORE_CASE = 4
    """Partial match. Case is ignored. Search criteria must be in the search result"""

    @staticmethod
    def from_str(s: str) -> "SearchMatchKind":
        """
        Gets an ``SearchMatchKind`` instance from string.

        Args:
            s (str): String that represents the name of an enum Name.
                ``s`` is case insensitive and can be ``CamelCase``, ``pascal_case`` , ``snake_case``,
                ``hyphen-case``, ``normal case``.

        Raises:
            ValueError: If input string is empty.
            AttributeError: If unable to get ``SearchMatchKind`` instance.

        Returns:
            SearchMatchKind: Enum instance.
        """
        return kind_helper.enum_from_string(s, SearchMatchKind)
