from enum import Enum
from ooodev.utils.kind import kind_helper


class LanguageKind(Enum):
    PYTHON = "Python"
    BASIC = "Basic"
    JAVA = "Java"
    JAVASCRIPT = "JavaScript"
    BEANS = "Beans"

    def __str__(self) -> str:
        return self.value

    @staticmethod
    def from_str(s: str) -> "LanguageKind":
        """
        Gets an ``LanguageKind`` instance from string.

        Args:
            s (str): String that represents the name of an enum Name.
                ``s`` is case insensitive and can be ``CamelCase``, ``pascal_case`` , ``snake_case``,
                ``hyphen-case``, ``normal case``.

        Raises:
            ValueError: If input string is empty.
            AttributeError: If unable to get ``DrawingHatchingKind`` instance.

        Returns:
            LanguageKind: Enum instance.
        """
        return kind_helper.enum_from_string(s, LanguageKind)
