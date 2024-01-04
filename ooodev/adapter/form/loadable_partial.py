from __future__ import annotations
from typing import TYPE_CHECKING

from com.sun.star.form import XLoadable

if TYPE_CHECKING:
    from com.sun.star.form import XLoadListener
    from ooodev.utils.type_var import UnoInterface


class LoadablePartial:
    """
    Partial Class for XLoadable.

    This interface does not really provide an own functionality, it is only for easier runtime identification of form components.
    """

    def __init__(self, component: XLoadable, interface: UnoInterface | None = XLoadable) -> None:
        """
        Constructor

        Args:
            component (XLoadable): UNO Component that implements ``com.sun.star.container.XLoadable``.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XLoadable``.
        """
        self.__component = component

    # region XLoadable
    def add_load_listener(self, listener: XLoadListener) -> None:
        """
        Adds a listener to be notified when the data is loaded.
        """
        self.__component.addLoadListener(listener)

    def is_loaded(self) -> bool:
        """
        Returns whether the data is loaded.
        """
        return self.__component.isLoaded()

    def load(self) -> None:
        """
        Loads the data.
        """
        self.__component.load()

    def reload(self) -> None:
        """
        Does a smart refresh of the object.
        """
        self.__component.reload()

    def remove_load_listener(self, listener: XLoadListener) -> None:
        """
        Removes a listener from the list of load listeners.
        """
        self.__component.removeLoadListener(listener)

    def unload(self) -> None:
        """
        Unloads the data.
        """
        self.__component.unload()

    # endregion XLoadable
