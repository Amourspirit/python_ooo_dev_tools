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

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

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
