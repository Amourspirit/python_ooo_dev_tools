from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.frame import XModuleManager

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.uno import XInterface
    from ooodev.utils.type_var import UnoInterface


class ModuleManagerPartial:
    """
    Partial class for XModuleManager.
    """

    def __init__(self, component: XModuleManager, interface: UnoInterface | None = XModuleManager) -> None:
        """
        Constructor

        Args:
            component (XModuleManager): UNO Component that implements ``com.sun.star.frame.XModuleManager`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XModuleManager``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XModuleManager
    def identify(self, module: XInterface) -> str:
        """
        Identifies the given module.

        This identifier can then be used at the service ModuleManager to get more information about this module.

        For identification the interface com.sun.star.lang.XServiceInfo is requested on the given module. Because all module service registrations must be unique this value can be queried and checked against the configuration.

        Since OOo 2.3.0 also the optional interface XModule will be used. If its exists it will be preferred.

        Raises:
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
            UnknownModuleException: ``UnknownModuleException``
        """
        return self.__component.identify(module)

    # endregion XModuleManager
