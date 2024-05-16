from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno
from com.sun.star.script import XStorageBasedLibraryContainer

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.adapter.script.persistent_library_container_partial import PersistentLibraryContainerPartial

if TYPE_CHECKING:
    from com.sun.star.embed import XStorage
    from ooodev.utils.type_var import UnoInterface


class StorageBasedLibraryContainerPartial(PersistentLibraryContainerPartial):
    """
    Partial class for XStorageBasedLibraryContainer.
    """

    def __init__(
        self, component: XStorageBasedLibraryContainer, interface: UnoInterface | None = XStorageBasedLibraryContainer
    ) -> None:
        """
        Constructor

        Args:
            component (XStorageBasedLibraryContainer): UNO Component that implements ``com.sun.star.script.XStorageBasedLibraryContainer`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XStorageBasedLibraryContainer``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        PersistentLibraryContainerPartial.__init__(self, component=component, interface=None)
        self.__component = component

    # region XStorageBasedLibraryContainer
    def store_libraries_to_storage(self, root_storage: XStorage) -> None:
        """
        Stores the libraries to a storage other than the current container storage

        Note that the library container is not automatically associated with the new root storage. Instead, you need to manually set the RootStorage attribute afterwards. This separation allows for Save-To as well Save-As operations.

        Raises:
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
            com.sun.star.lang.WrappedTargetException: ``WrappedTargetException``
        """
        self.__component.storeLibrariesToStorage(root_storage)

    @property
    def root_storage(self) -> XStorage:
        """
        Denotes the root storage associated with the container.

        Effectively, this attribute is a typed version of XPersistentLibraryContainer.RootLocation, it's guaranteed that at every time, XPersistentLibraryContainer.RootLocation and root_storage have the same value.

        You should only set this attribute to a new value if you previously called storeLibrariesToStorage with the same storage. Setting this attribute to a storage into which the container has not been stored previously might result in unexpected behavior.
        """
        return self.__component.RootStorage

    @root_storage.setter
    def root_storage(self, value: XStorage) -> None:
        self.__component.RootStorage = value

    # endregion XStorageBasedLibraryContainer
