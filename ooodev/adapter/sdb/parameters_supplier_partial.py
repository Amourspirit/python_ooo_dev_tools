from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.sdb import XParametersSupplier

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

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

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XParametersSupplier
    def get_parameters(self) -> XIndexAccess:
        """
        Returns the container of parameters.

        Returns:
            XIndexAccess: The parameters.
        """
        return self.__component.getParameters()

    # endregion XParametersSupplier
