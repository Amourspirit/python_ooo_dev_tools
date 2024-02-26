from enum import Enum
from ooodev.utils.kind import kind_helper

# these are helper lookups


class InfoPathsKind(str, Enum):
    """
    Info Path Lookup

    See Also:
        :py:meth:`.Info.get_paths`
    """

    ADDIN = "Addin"
    AUTO_CORRECT = "AutoCorrect"
    AUTO_TEXT = "AutoText"
    BACKUP = "Backup"
    BASIC = "Basic"
    BITMAP = "Bitmap"
    CONFIG = "Config"
    DICTIONARY = "Dictionary"
    FAVORITE = "Favorite"
    FILTER = "Filter"
    GALLERY = "Gallery"
    GRAPHIC = "Graphic"
    HELP = "Help"
    LINGUISTIC = "Linguistic"
    MODULE = "Module"
    PALETTE = "Palette"
    PLUGIN = "Plugin"
    STORAGE = "Storage"
    TEMP = "Temp"
    TEMPLATE = "Template"
    UI_CONFIG = "UIConfig"
    USER_CONFIG = "UserConfig"
    USER_DICTIONARY = "UserDictionary"
    """UserDictionary (deprecated)"""
    WORK = "Work"

    def __str__(self) -> str:
        return self.value

    @staticmethod
    def from_str(s: str) -> "InfoPathsKind":
        """
        Gets an ``InfoPathsKind`` instance from string.

        Args:
            s (str): String that represents the name of an enum Name.
                ``s`` is case insensitive and can be ``CamelCase``, ``pascal_case`` , ``snake_case``,
                ``hyphen-case``, ``normal case``.

        Raises:
            ValueError: If input string is empty.
            AttributeError: If unable to get ``InfoPathsKind`` instance.

        Returns:
            InfoPathsKind: Enum instance.
        """
        return kind_helper.enum_from_string(s, InfoPathsKind)
