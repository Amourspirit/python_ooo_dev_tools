from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.lang import XInitialization

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class InitializationPartial:
    """
    Partial class for XInitialization.
    """

    def __init__(self, component: XInitialization, interface: UnoInterface | None = XInitialization) -> None:
        """
        Constructor

        Args:
            component (XInitialization): UNO Component that implements ``com.sun.star.lang.XInitialization`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XInitialization``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XInitialization
    def initialize(self, *args: Any) -> None:
        """
        Initializes the object.

        It should be called directly after the object is created.

        Raises:
            com.sun.star.uno.Exception: ``Exception``
        """
        self.__component.initialize(args)

    # endregion XInitialization
