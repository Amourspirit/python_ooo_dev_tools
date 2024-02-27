from __future__ import annotations
from typing import TYPE_CHECKING

import uno
from com.sun.star.frame import XStorable2

from ooodev.adapter.frame.storable_partial import StorablePartial

if TYPE_CHECKING:
    from com.sun.star.beans import PropertyValue
    from ooodev.utils.type_var import UnoInterface


class Storable2Partial(StorablePartial):
    """
    Partial class for XStorable2.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XStorable2, interface: UnoInterface | None = XStorable2) -> None:
        """
        Constructor

        Args:
            component (XStorable2): UNO Component that implements ``com.sun.star.frame.XStorable2``.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XStorable2``.
        """
        StorablePartial.__init__(self, component, interface)
        self.__component = component

    # region XStorable2
    def store_self(self, *args: PropertyValue) -> None:
        """
        Stores the data to the URL from which it was loaded.
        Only objects which know their locations can be stored.
        This is an extension of the ``XStorable.store()``.
        This method allows to specify some additional parameters for storing process.

        Args:
            *args (PropertyValue): Additional parameters for storing process.

        Raises:
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
            com.sun.star.io.IOException: ``IOException``
        """
        self.__component.storeSelf(args)

    # endregion XStorable2
