from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno
from com.sun.star.io import XPersistObject

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface
    from com.sun.star.io import XObjectInputStream
    from com.sun.star.io import XObjectOutputStream


class PersistObjectPartial:
    """
    Partial Class XPersistObject.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XPersistObject, interface: UnoInterface | None = XPersistObject) -> None:
        """
        Constructor

        Args:
            component (XPersistObject): UNO Component that implements ``com.sun.star.io.XPersistObject`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XPersistObject``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XPersistObject
    def get_service_name(self) -> str:
        """
        Returns the service name of the object.

        Returns:
            str: The service name.
        """
        return self.__component.getServiceName()

    def read(self, stream: XObjectInputStream) -> None:
        """
        Reads the object from the given stream.

        Args:
            stream (XObjectInputStream): The stream to read from.
        """
        self.__component.read(stream)

    def write(self, stream: XObjectOutputStream) -> None:
        """
        Writes the object to the given stream.

        Args:
            stream (XObjectOutputStream): The stream to write to.
        """
        self.__component.write(stream)

    # endregion XPersistObject
