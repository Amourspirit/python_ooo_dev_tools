from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno
from com.sun.star.script import XLibraryContainer2

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.adapter.script.library_container_partial import LibraryContainerPartial

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class LibraryContainer2Partial(LibraryContainerPartial):
    """
    Partial class for XLibraryContainer2.
    """

    def __init__(self, component: XLibraryContainer2, interface: UnoInterface | None = XLibraryContainer2) -> None:
        """
        Constructor

        Args:
            component (XLibraryContainer2): UNO Component that implements ``com.sun.star.script.XLibraryContainer2`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XLibraryContainer2``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        LibraryContainerPartial.__init__(self, component=component, interface=None)
        self.__component = component

    # region XLibraryContainer2
    def get_library_link_url(self, name: str) -> str:
        """
        Returns the location of the library link target.

        Should return the same URL that was passed to createLibraryLink in the StorageURL parameter.

        If the accessed library item exists but isn't a link, an IllegalArgumentException is thrown

        Raises:
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
            com.sun.star.container.NoSuchElementException: ``NoSuchElementException``
        """
        return self.__component.getLibraryLinkURL(name)

    def is_library_link(self, name: str) -> bool:
        """
        returns true if the accessed library item is a link, e.g., created by createLibraryLink, otherwise false.

        Raises:
            com.sun.star.container.NoSuchElementException: ``NoSuchElementException``
        """
        return self.__component.isLibraryLink(name)

    def is_library_read_only(self, name: str) -> bool:
        """
        Returns true if the accessed library item (library or library link) is read only.

        A library can be read only because it was set to read only using the methods provided by this interface or because of other reasons depending on the implementation (e.g., file system write protection)

        Raises:
            com.sun.star.container.NoSuchElementException: ``NoSuchElementException``
        """
        return self.__component.isLibraryReadOnly(name)

    def rename_library(self, name: str, new_name: str) -> None:
        """
        Renames the library item with the specified name.

        If the accessed library item is a link only the link is renamed, not the target library. If a library with the new name exists already a com.sun.star.container.ElementExistException is thrown.

        Raises:
            com.sun.star.container.NoSuchElementException: ``NoSuchElementException``
            com.sun.star.container.ElementExistException: ``ElementExistException``
        """
        return self.__component.renameLibrary(name, new_name)

    def set_library_read_only(self, name: str, read_only: bool) -> None:
        """
        Sets the accessed library item (library or library link) to read only according to the flag bReadOnly (true means read only)

        Raises:
            com.sun.star.container.NoSuchElementException: ``NoSuchElementException``
        """
        return self.__component.setLibraryReadOnly(name, read_only)

    # endregion XLibraryContainer2
