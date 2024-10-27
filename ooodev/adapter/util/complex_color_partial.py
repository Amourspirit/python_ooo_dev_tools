from __future__ import annotations
from typing import Any, TYPE_CHECKING

# by not importing XComplexColor, we can avoid causing issues with pre 7.6 versions of LibreOffice
# from com.sun.star.util import XComplexColor  # type: ignore

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface
    from com.sun.star.util import XTheme  # type: ignore


class ComplexColorPartial:
    """
    Partial Class XComplexColor.

    Since LibreOffice 7.6, the interface ``com.sun.star.util.XComplexColor`` is available.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: Any, interface: UnoInterface | None = Any) -> None:
        """
        Constructor

        Args:
            component (XComplexColor): UNO Component that implements ``com.sun.star.util.XComplexColor`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XComplexColor``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XTheme
    def get_theme_color_type(self) -> int:
        """
        Type of the theme color.

        Returns:
            str: The clone.
        """
        return self.__component.getThemeColorType()

    def get_type(self) -> int:
        """
        Type of the color.

        Returns:
            int: The color type.
        """
        return self.__component.getType()

    def resolve_color(self, theme: XTheme) -> int:
        """
        Type of the color.

        Returns:
            int: The color type.
        """
        return self.__component.resolveColor(theme)

    # endregion XTheme
