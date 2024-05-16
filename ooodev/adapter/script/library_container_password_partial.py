from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno
from com.sun.star.script import XLibraryContainerPassword

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class LibraryContainerPasswordPartial:
    """
    Partial class for XLibraryContainerPassword.
    """

    def __init__(
        self, component: XLibraryContainerPassword, interface: UnoInterface | None = XLibraryContainerPassword
    ) -> None:
        """
        Constructor

        Args:
            component (XLibraryContainerPassword): UNO Component that implements ``com.sun.star.script.XLibraryContainerPassword`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XLibraryContainerPassword``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XLibraryContainerPassword
    def change_library_password(self, name: str, old_password: str, new_password: str) -> None:
        """
        Changes the library's password.

        If the library wasn't password protected before: The OldPassword parameter has to be an empty string. Afterwards calls to isLibraryPasswordProtected and isLibraryPasswordVerified for this library will return true.

        If the library already was password protected: The OldPassword parameter has to be set to the previous defined password. If then the NewPassword parameter is an empty string the library password protection will be disabled afterwards (afterwards calls to isLibraryPasswordProtected for this library will return false). If the NewPassword parameter is not an empty string it will accepted as the new password for the library.

        If a library with the this name doesn't exist but isn't com.sun.star.container.NoSuchElementException is thrown.

        If the library exists and is password protected and a wrong OldPassword is passed to the method a com.sun.star.lang.IllegalArgumentException is thrown.

        If the library exists and isn't password protected and the OldPassword isn't an empty string or the library is read only a com.sun.star.lang.IllegalArgumentException is thrown.

        Raises:
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
            com.sun.star.container.NoSuchElementException: ``NoSuchElementException``
        """
        self.__component.changeLibraryPassword(name, old_password, new_password)

    def is_library_password_protected(self, name: str) -> bool:
        """
        Returns true if the accessed library item is protected by a password.

        If a library with the this name doesn't exist a com.sun.star.container.NoSuchElementException is thrown.

        Raises:
            com.sun.star.container.NoSuchElementException: ``NoSuchElementException``
        """
        return self.__component.isLibraryPasswordProtected(name)

    def is_library_password_verified(self, name: str) -> bool:
        """
        Returns true if the accessed library item is protected by a password (see isLibraryPasswordProtected) and the password was already verified with verifyLibraryPassword or if an initial password was set with changeLibraryPassword.

        If a library with the this name doesn't exist a com.sun.star.container.NoSuchElementException is thrown.

        If the library exists but isn't password protected a com.sun.star.lang.IllegalArgumentException is thrown.

        Raises:
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
            com.sun.star.container.NoSuchElementException: ``NoSuchElementException``
        """
        return self.__component.isLibraryPasswordVerified(name)

    def verify_library_password(self, name: str, password: str) -> bool:
        """
        Verifies the library's password.

        If the correct password was passed, the method returns true and further calls to isLibraryPasswordVerified will also return true.

        If a library with the this name doesn't exist a com.sun.star.container.NoSuchElementException is thrown.

        If the library exists but isn't password protected a com.sun.star.lang.IllegalArgumentException is thrown.

        If the library password is already verified a com.sun.star.lang.IllegalArgumentException is thrown.

        Raises:
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
            com.sun.star.container.NoSuchElementException: ``NoSuchElementException``
        """
        return self.__component.verifyLibraryPassword(name, password)

    # endregion XLibraryContainerPassword
