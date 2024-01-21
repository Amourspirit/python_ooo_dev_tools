from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.document import XFilter

from ooodev.exceptions import ex as mEx
from ooodev.utils import lo as mLo


if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface
    from com.sun.star.beans import PropertyValue


class FilterPartial:
    """
    Partial class for XFilter.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XFilter, interface: UnoInterface | None = XFilter) -> None:
        """
        Constructor

        Args:
            component (XFilter): UNO Component that implements ``com.sun.star.container.XFilter`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XFilter``.
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

    # region XFilter
    def cancel(self) -> None:
        """Cancels the export process."""
        self.__component.cancel()

    def filter(self, *args: PropertyValue) -> None:
        """
        Filter the document.

        Args:
            url (str): The URL of the graphic file to be exported.
            args (Tuple[PropertyValue, ...]): The arguments to be used for the export filter.
        """
        # desc = uno.Any("[]com.sun.star.beans.PropertyValue", args)  # type: ignore
        # uno.invoke(self.__component, "filter", (desc,))
        self.__component.filter(args)

    # endregion XFilter
