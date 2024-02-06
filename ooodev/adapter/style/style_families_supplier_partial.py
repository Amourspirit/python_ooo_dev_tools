from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.style import XStyleFamiliesSupplier

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.utils.type_var import UnoInterface

if TYPE_CHECKING:
    from com.sun.star.container import XNameAccess


class StyleFamiliesSupplierPartial:
    """
    Partial class for XStyleFamiliesSupplier.
    """

    def __init__(
        self, component: XStyleFamiliesSupplier, interface: UnoInterface | None = XStyleFamiliesSupplier
    ) -> None:
        """
        Constructor

        Args:
            component (XStyleFamiliesSupplier): UNO Component that implements ``com.sun.star.style.XStyleFamiliesSupplier`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XStyleFamiliesSupplier``.
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

    # region XStyleFamiliesSupplier
    def get_style_families(self) -> XNameAccess:
        """
        This method returns the collection of style families available in the container document.
        """
        return self.__component.getStyleFamilies()

    # endregion XStyleFamiliesSupplier
