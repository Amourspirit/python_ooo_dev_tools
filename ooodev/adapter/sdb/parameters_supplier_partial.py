from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.sdb import XParametersSupplier

from ooodev.exceptions import ex as mEx
from ooodev.utils import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.container import XIndexAccess
    from ooodev.utils.type_var import UnoInterface


class ParametersSupplierPartial:
    """
    Partial class for XParametersSupplier.
    """

    def __init__(self, component: XParametersSupplier, interface: UnoInterface | None = XParametersSupplier) -> None:
        """
        Constructor

        Args:
            component (XParametersSupplier): UNO Component that implements ``com.sun.star.container.XParametersSupplier`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XParametersSupplier``.
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

    # region XParametersSupplier
    def get_parameters(self) -> XIndexAccess:
        """
        Returns the container of parameters.

        Returns:
            XIndexAccess: The parameters.
        """
        return self.__component.getParameters()

    # endregion XParametersSupplier
