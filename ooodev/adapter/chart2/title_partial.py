from __future__ import annotations
from typing import Any, TYPE_CHECKING, Tuple
import uno

from com.sun.star.chart2 import XTitle

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.chart2 import XFormattedString
    from ooodev.utils.type_var import UnoInterface


class TitlePartial:
    """
    Partial class for XTitle.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XTitle, interface: UnoInterface | None = XTitle) -> None:
        """
        Constructor

        Args:
            component (XTitle): UNO Component that implements ``com.sun.star.chart2.XTitle`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XTitle``.
        """
        self.__interface = interface
        self.__validate(component)
        self.__component = component

    def __validate(self, component: Any) -> None:
        """
        Validates the component.

        Args:
            component (Any): The component to be validated.
        """
        if self.__interface is None:
            return
        if not mLo.Lo.is_uno_interfaces(component, self.__interface):
            raise mEx.MissingInterfaceError(self.__interface)

    # region XTitle
    def get_text(self) -> Tuple[XFormattedString, ...]:
        """
        Gets the text of the title.
        """
        return self.__component.getText()

    def set_text(self, *strings: XFormattedString) -> None:
        """
        Sets the text of the title.
        """
        self.__component.setText(strings)

    # endregion XTitle
