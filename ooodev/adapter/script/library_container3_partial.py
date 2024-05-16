from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno
from com.sun.star.script import XLibraryContainer3

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.adapter.script.library_container2_partial import LibraryContainer2Partial

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class LibraryContainer3Partial(LibraryContainer2Partial):
    """
    Partial class for XLibraryContainer3.
    """

    def __init__(self, component: XLibraryContainer3, interface: UnoInterface | None = XLibraryContainer3) -> None:
        """
        Constructor

        Args:
            component (XLibraryContainer3): UNO Component that implements ``com.sun.star.script.XLibraryContainer3`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XLibraryContainer3``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        LibraryContainer2Partial.__init__(self, component=component, interface=None)
        self.__component = component

    # region XLibraryContainer3
    def get_original_library_link_url(self, name: str) -> str:
        """
        Returns the location of the library link target.

        The returned URL is literally the same as the one provided in ``create_library_link()``.
        Otherwise the behavior is the same as ``get_library_link_url()``

        Raises:
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
            com.sun.star.container.NoSuchElementException: ``NoSuchElementException``
        """
        return self.__component.getOriginalLibraryLinkURL(name)

    # endregion XLibraryContainer3
