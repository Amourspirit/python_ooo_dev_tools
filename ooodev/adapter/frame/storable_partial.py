from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.frame import XStorable

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.beans import PropertyValue
    from ooodev.utils.type_var import UnoInterface


class StorablePartial:
    """
    Partial class for XStorable.
    """

    def __init__(self, component: XStorable, interface: UnoInterface | None = XStorable) -> None:
        """
        Constructor

        Args:
            component (XStorable): UNO Component that implements ``com.sun.star.frame.XStorable`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XStorable``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XStorable
    def get_location(self) -> str:
        """
        After XStorable.storeAsURL() it returns the URL the object was stored to.
        """
        return self.__component.getLocation()

    def has_location(self) -> bool:
        """
        The object may know the location because it was loaded from there, or because it is stored there.
        """
        return self.__component.hasLocation()

    def is_readonly(self) -> bool:
        """
        It is not possible to call XStorable.store() successfully when the data store is read-only.
        """
        return self.__component.isReadonly()

    def store(self) -> None:
        """
        Stores the data to the URL from which it was loaded.
        Only objects which know their locations can be stored.

        Raises:
            com.sun.star.io.IOException: ``IOException``
        """
        self.__component.store()

    def store_as_url(self, url: str, *args: PropertyValue) -> None:
        """
        Stores the object's persistent data to a URL and makes this URL the new location of the object.
        This is the normal behavior for UI ``save-as`` feature.
        The change of the location makes it necessary to store the document in a format that the object can load. For this reason the implementation of XStorable.storeAsURL() will throw an exception if a pure export filter is used, it will accept only combined import/export filters. For such filters the method XStorable.storeToURL() must be used that does not change the location of the object.

        Args:
            url (str): The URL to be stored.
            *args (PropertyValue): Additional parameters for storing process.

        Raises:
            com.sun.star.io.IOException: ``IOException``
        """
        self.__component.storeAsURL(url, args)

    def store_to_url(self, url: str, *args: PropertyValue) -> None:
        """
        Stores the object's persistent data to a URL and continues to be a representation of the old URL.
        This is the normal behavior for UI export feature.
        This method accepts all kinds of export filters, not only combined import/export filters because it implements an exporting capability, not a persistence capability.

        Args:
            url (str): The URL to be stored.
            *args (PropertyValue): Additional parameters for storing process.

        Raises:
            com.sun.star.io.IOException: ``IOException``
        """
        self.__component.storeToURL(url, args)

    # endregion XStorable
