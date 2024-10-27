from __future__ import annotations
from typing import Dict, TYPE_CHECKING, Tuple

if TYPE_CHECKING:
    from com.sun.star.util import XTheme  # type: ignore


class ThemePropertiesPartial:
    """
    Partial Properties class for XTheme Service.
    """

    def __init__(self, component: XTheme) -> None:
        """
        Constructor

        Args:
            component (XTheme): UNO Component that implements ``com.sun.star.text.XTheme`` interface.
        """
        self.__component = component

    # region Properties

    @property
    def color_set(self) -> Tuple[int, int, int, int, int, int, int, int, int, int, int, int]:
        """
        Gets the color set of the theme.

        The color set is a sequence of 12 colors: Dark 1, Light 1, Dark 2, Light 2, Accent 1, Accent 2, Accent 3, Accent 4, Accent 5, Accent 6, Hyperlink, FollowedHyperlink

        Returns:
            tuple: The color set.
        """
        return self.get_color_set()  # type: ignore

    @property
    def name(self) -> str:
        """
        Gets the name of the theme.

        Returns:
            str: The name.
        """
        return self.get_name()  # type: ignore

    @property
    def color_scheme_name(self) -> str:
        """
        Gets the color scheme name of the theme.

        Returns:
            str: The color scheme name. Maybe empty string.
        """
        return self.get_color_scheme_name()  # type: ignore

    @property
    def color_set_dict(self) -> Dict[str, int]:
        """
        Gets a dictionary mapping color names to their corresponding integer values.

        The color names include:

        - "dark1"
        - "light1"
        - "dark2"
        - "light2"
        - "accent1"
        - "accent2"
        - "accent3"
        - "accent4"
        - "accent5"
        - "accent6"
        - "hyperlink"
        - "followed_hyperlink"

        Returns:
            Dict[str, int]: A dictionary where the keys are color names and the values are integers representing the colors.
        """
        # get_color_set_dict() is in theme_partial.py
        return self.get_color_set_dict()  # type: ignore

    # endregion Properties
