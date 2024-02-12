from __future__ import annotations
from typing import TYPE_CHECKING

from com.sun.star.style import XStyle
from ooodev.adapter.container.named_partial import NamedPartial

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class StylePartial(NamedPartial):
    """
    Partial Class for XStyle.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XStyle, interface: UnoInterface | None = XStyle) -> None:
        """
        Constructor

        Args:
            component (XStyle): UNO Component that implements ``com.sun.star.container.XStyle``.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XStyle``.
        """
        NamedPartial.__init__(self, component, interface)
        self.__component = component

    # region XStyle
    def is_user_defined(self) -> bool:
        """Returns ``True`` if this style is user defined."""
        return self.__component.isUserDefined()

    def is_in_use(self) -> bool:
        """Returns ``True`` if this style is in use."""
        return self.__component.isInUse()

    def get_parent_style(self) -> str:
        """Returns the name of the parent style."""
        return self.__component.getParentStyle()

    def set_parent_style(self, parent_style: str) -> None:
        """
        Sets the name of the parent style.

        Args:
            parent_style (str): The name of the parent style.
        """
        self.__component.setParentStyle(parent_style)

    # endregion XStyle
