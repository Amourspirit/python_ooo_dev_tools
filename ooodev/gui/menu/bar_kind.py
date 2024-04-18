from enum import Enum
from ooodev.utils.kind import kind_helper


class BarKind(Enum):
    """Specifies the horizontal alignment of the text in the control."""

    MENU_BAR = "private:resource/menubar/menubar"
    STATUS_BAR = "private:resource/statusbar/statusbar"
    FIND_BAR = "private:resource/toolbar/findbar"
    STANDARD_BAR = "private:resource/toolbar/standardbar"
    TOOL_BAR = "private:resource/toolbar/toolbar"

    def __str__(self) -> str:
        return self.value

    @staticmethod
    def from_str(s: str) -> "BarKind":
        """
        Gets an ``BarKind`` instance from string.

        Args:
            s (str): String that represents the name of an enum Name.

        Raises:
            ValueError: If input string is empty.
            AttributeError: If unable to get ``BarKind`` instance.

        Returns:
            BarKind: Enum instance.
        """
        return kind_helper.enum_from_string(s, BarKind)
