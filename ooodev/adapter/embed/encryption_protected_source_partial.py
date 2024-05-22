from __future__ import annotations
from typing import Any
import uno

from com.sun.star.embed import XEncryptionProtectedSource

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.utils.type_var import UnoInterface


class EncryptionProtectedSourcePartial:
    """
    Partial class for XEncryptionProtectedSource.
    """

    def __init__(
        self, component: XEncryptionProtectedSource, interface: UnoInterface | None = XEncryptionProtectedSource
    ) -> None:
        """
        Constructor

        Args:
            component (XEncryptionProtectedSource): UNO Component that implements ``com.sun.star.embed.XEncryptionProtectedSource`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XEncryptionProtectedSource``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XEncryptionProtectedSource
    def remove_encryption(self) -> None:
        """
        removes encryption from the object.

        Raises:
            com.sun.star.io.IOException: ``IOException``
        """
        self.__component.removeEncryption()

    def set_encryption_password(self, password: str) -> None:
        """
        sets a password for the object.

        Raises:
            com.sun.star.io.IOException: ``IOException``
        """
        self.__component.setEncryptionPassword(password)

    # endregion XEncryptionProtectedSource
