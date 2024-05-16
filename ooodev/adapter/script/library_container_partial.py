from __future__ import annotations
from typing import Any, overload, TYPE_CHECKING
import uno
from com.sun.star.script import XLibraryContainer

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.adapter.container.name_container_comp import NameContainerComp
from ooodev.adapter.container.name_access_comp import NameAccessComp

if TYPE_CHECKING:
    from typing_extensions import Literal
    from ooodev.utils.type_var import UnoInterface


class LibraryContainerPartial:
    """
    Partial class for XLibraryContainer.
    """

    def __init__(self, component: XLibraryContainer, interface: UnoInterface | None = XLibraryContainer) -> None:
        """
        Constructor

        Args:
            component (XLibraryContainer): UNO Component that implements ``com.sun.star.script.XLibraryContainer`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XLibraryContainer``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XLibraryContainer
    def create_library(self, name: str) -> NameContainerComp:
        """
        Creates a new library.

        Args:
            name (str): The name of the library.

        Raises:
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
            com.sun.star.container.ElementExistException: ``ElementExistException``
        """
        return NameContainerComp(self.__component.createLibrary(name))

    @overload
    def create_library_link(self, name: str, storage_url: str, read_only: Literal[True]) -> NameAccessComp: ...

    @overload
    def create_library_link(self, name: str, storage_url: str, read_only: Literal[False]) -> NameContainerComp: ...

    def create_library_link(
        self, name: str, storage_url: str, read_only: bool = False
    ) -> NameAccessComp | NameContainerComp:
        """
        Creates a link to an ``external`` library that then can be accessed using this library manager.

        The format of the StorageURL is a matter of the implementation.
        If the read only flag is set, the linked library cannot be modified.
        In this case, the returned component really is only an NameAccessComp.
        If ``read_only`` is ``False`` (and the referenced storage allows write access) the returned component also can be a
        NameContainerComp.

        Args:
            name (str): The name of the library link.
            storage_url (str): The URL of the library to be linked.
            read_only (bool): If true, the linked library cannot be modified.

        Raises:
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
            com.sun.star.container.ElementExistException: ``ElementExistException``

        Returns:
            NameAccessComp: The created library link.
        """
        result = self.__component.createLibraryLink(name, storage_url, read_only)
        if mLo.Lo.is_uno_interfaces(result, "com.sun.star.container.XNameContainer"):
            return NameContainerComp(result)  # type: ignore
        return NameAccessComp(result)

    def is_library_loaded(self, name: str) -> bool:
        """
        returns true if the accessed library is already loaded from the storage, otherwise false.

        Raises:
            com.sun.star.container.NoSuchElementException: ``NoSuchElementException``
        """
        return self.__component.isLibraryLoaded(name)

    def load_library(self, name: str) -> None:
        """
        Causes the accessed library to be loaded from its storage if it hasn't already been loaded.

        Raises:
            com.sun.star.container.NoSuchElementException: ``NoSuchElementException``
            com.sun.star.lang.WrappedTargetException: ``WrappedTargetException``
        """
        self.__component.loadLibrary(name)

    def remove_library(self, name: str) -> None:
        """
        removes the library item with the specified name.

        If the accessed library item is a link only the link is removed, not the target library.

        Raises:
            com.sun.star.container.NoSuchElementException: ``NoSuchElementException``
            com.sun.star.lang.WrappedTargetException: ``WrappedTargetException``
        """
        self.__component.removeLibrary(name)

    # endregion XLibraryContainer
