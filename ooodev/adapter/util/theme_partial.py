from __future__ import annotations
from typing import Any, Dict, Tuple, TYPE_CHECKING

# by not importing XTheme, we can avoid causing issues with pre 7.6 versions of LibreOffice
# from com.sun.star.util import XTheme  # type: ignore

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.utils import props as mProps
from ooodev.utils import info as mInfo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class ThemePartial:
    """
    Partial Class XTheme.

    Since LibreOffice 7.6, the interface ``com.sun.star.util.XTheme`` is available.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: Any, interface: UnoInterface | None = Any) -> None:
        """
        Constructor

        Args:
            component (XTheme, tuple): UNO Component that implements ``com.sun.star.util.XTheme`` interface or Tuple of property value.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XTheme``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XTheme
    def get_name(self) -> str:
        """
        Gets the name of the theme.

        Returns:
            str: The name.

        Note:
            If the component is a tuple of property values, the method will return the value of the ``Name`` property.
            If the ``Name`` property is not found, the method will return an empty string.
        """
        if hasattr(self.__component, "getName"):
            return self.__component.getName()
        # assume tuple of property values
        prop_dict = mProps.Props.data_to_dict(self.__component)
        return prop_dict.get("Name", "")

    def get_color_set(self) -> Tuple[int, int, int, int, int, int, int, int, int, int, int, int]:
        """
        Gets the color set of the theme.

        The color set is a sequence of 12 colors: Dark 1, Light 1, Dark 2, Light 2, Accent 1, Accent 2, Accent 3, Accent 4, Accent 5, Accent 6, Hyperlink, FollowedHyperlink

        Returns:
            tuple: The color set. Maybe a tuple of ``-1`` values. If any tuple value is ``-1`` then the color was not able to be retrieved.

        Note:
            If the component is a tuple of property values, the method will return the value of the ``ColorScheme`` property.
            If the ``ColorScheme`` property is not found, the method will return a tuple of -1 values.
        """
        if hasattr(self.__component, "getColorSet"):
            return self.__component.getColorSet()

        prop_dict = mProps.Props.data_to_dict(self.__component)
        return prop_dict.get("ColorScheme", (-1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1))

    # endregion XTheme

    def get_color_scheme_name(self) -> str:
        """
        Gets the color scheme name of the theme.

        Returns:
            str: The color scheme name. Maybe empty string.

        Note:
            If the component is a tuple of property values, the method will return the value of the ``ColorSchemeName`` property.
            If the ``ColorSchemeName`` property is not found, the method will return an empty string.
        """
        if not mInfo.Info.is_instance(self.__component, tuple):
            return ""

        try:
            prop_dict = mProps.Props.data_to_dict(self.__component)
            return prop_dict.get("ColorSchemeName", "")
        except Exception:
            return ""
        # assume tuple of property values

    def get_color_set_dict(self) -> Dict[str, int]:
        """
        Returns a dictionary mapping color names to their corresponding integer values.

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

        names = (
            "dark1",
            "light1",
            "dark2",
            "light2",
            "accent1",
            "accent2",
            "accent3",
            "accent4",
            "accent5",
            "accent6",
            "hyperlink",
            "followed_hyperlink",
        )
        return dict(zip(names, self.get_color_set()))
