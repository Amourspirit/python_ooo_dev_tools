from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.document import XStorageBasedDocument

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo


if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface
    from com.sun.star.beans import PropertyValue
    from com.sun.star.document import XStorageChangeListener
    from com.sun.star.embed import XStorage


class StorageBasedDocumentPartial:
    """
    Partial class for XStorageBasedDocument.
    """

    def __init__(
        self, component: XStorageBasedDocument, interface: UnoInterface | None = XStorageBasedDocument
    ) -> None:
        """
        Constructor

        Args:
            component (XStorageBasedDocument): UNO Component that implements ``com.sun.star.document.XStorageBasedDocument`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XStorageBasedDocument``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XStorageBasedDocument
    def add_storage_change_listener(self, listener: XStorageChangeListener) -> None:
        """
        Allows to register a listener that will be notified when another storage is set to the document.
        """
        self.__component.addStorageChangeListener(listener)

    def get_document_storage(self) -> XStorage:
        """
        Allows to get the storage the document is based on.

        Raises:
            com.sun.star.io.IOException: ``IOException``
            com.sun.star.uno.Exception: ``Exception``
        """
        return self.__component.getDocumentStorage()

    def load_from_storage(self, storage: XStorage, *media_descriptor: PropertyValue) -> None:
        """
        Lets the document load itself using provided storage.

        Raises:
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
            com.sun.star.frame.DoubleInitializationException: ``DoubleInitializationException``
            com.sun.star.io.IOException: ``IOException``
            com.sun.star.uno.Exception: ``Exception``
        """
        self.__component.loadFromStorage(storage, media_descriptor)

    def remove_storage_change_listener(self, listener: XStorageChangeListener) -> None:
        """
        Allows to deregister the listener.
        """
        self.__component.removeStorageChangeListener(listener)

    def store_to_storage(self, storage: XStorage, *media_descriptor: PropertyValue) -> None:
        """
        Lets the document store itself to the provided storage.

        Raises:
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
            com.sun.star.io.IOException: ``IOException``
            com.sun.star.uno.Exception: ``Exception``
        """
        self.__component.storeToStorage(storage, media_descriptor)

    def switch_to_storage(self, storage: XStorage) -> None:
        """
        Allows to switch the document to the provided storage.

        Raises:
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
            com.sun.star.io.IOException: ``IOException``
            com.sun.star.uno.Exception: ``Exception``
        """
        self.__component.switchToStorage(storage)

    # endregion XStorageBasedDocument
