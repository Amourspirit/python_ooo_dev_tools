from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.util import XNumberFormatsSupplier
from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface
    from com.sun.star.beans import XPropertySet
    from com.sun.star.util import XNumberFormats


class NumberFormatsSupplierPartial:
    """
    Partial class for XNumberFormatsSupplier.
    """

    def __init__(
        self, component: XNumberFormatsSupplier, interface: UnoInterface | None = XNumberFormatsSupplier
    ) -> None:
        """
        Constructor

        Args:
            component (XNumberFormatsSupplier): UNO Component that implements ``com.sun.star.util.XNumberFormatsSupplier`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XNumberFormatsSupplier``.
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

    # region XNumberFormatsSupplier
    def get_number_format_settings(self) -> XPropertySet:
        """
        Gets the number format settings.
        """
        return self.__component.getNumberFormatSettings()

    def get_number_formats(self) -> XNumberFormats:
        """
        Gets the number formats.
        """
        return self.__component.getNumberFormats()

    # endregion XNumberFormatsSupplier
