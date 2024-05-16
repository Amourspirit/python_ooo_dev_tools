from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno
from com.sun.star.script import XPersistentLibraryContainer

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.adapter.script.library_container2_partial import LibraryContainer2Partial
from ooodev.adapter.util.modifiable_partial import ModifiablePartial

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class PersistentLibraryContainerPartial(ModifiablePartial, LibraryContainer2Partial):
    """
    Partial class for XPersistentLibraryContainer.
    """

    def __init__(
        self, component: XPersistentLibraryContainer, interface: UnoInterface | None = XPersistentLibraryContainer
    ) -> None:
        """
        Constructor

        Args:
            component (XPersistentLibraryContainer): UNO Component that implements ``com.sun.star.script.XPersistentLibraryContainer`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XPersistentLibraryContainer``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        ModifiablePartial.__init__(self, component=component, interface=None)
        LibraryContainer2Partial.__init__(self, component=component, interface=None)
        self.__component = component

    # region XPersistentLibraryContainer
    def store_libraries(self) -> None:
        """
        Stores the libraries to the current location.

        Raises:
            com.sun.star.lang.WrappedTargetException: ``WrappedTargetException``
        """
        self.__component.storeLibraries()

    @property
    def container_location_name(self) -> str:
        """
        Denotes the name of the sub location where the container elements are actually stored.
        """
        return self.__component.ContainerLocationName

    @container_location_name.setter
    def container_location_name(self, value: str) -> None:
        self.__component.ContainerLocationName = value

    @property
    def root_location(self) -> Any:
        """
        Denotes the root location associated with the container.

        The type of this location - it might be a folder in a file system, a storage, or anything else - is not specified here,
        but in derived interfaces or services implementing ``XPersistentLibraryContainer``.

        All operations of the library container take place in a location below the root location,
        the so-called container location, whose name is exposed as ``container_location_name``.
        """
        return self.__component.RootLocation

    @root_location.setter
    def root_location(self, value: Any) -> None:
        self.__component.RootLocation = value

    # endregion XPersistentLibraryContainer
