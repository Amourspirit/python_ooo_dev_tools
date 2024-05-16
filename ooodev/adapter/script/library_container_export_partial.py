from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno
from com.sun.star.script import XLibraryContainerExport

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface
    from com.sun.star.task import XInteractionHandler


class LibraryContainerExportPartial:
    """
    Partial class for XLibraryContainerExport.
    """

    def __init__(
        self, component: XLibraryContainerExport, interface: UnoInterface | None = XLibraryContainerExport
    ) -> None:
        """
        Constructor

        Args:
            component (XLibraryContainerExport): UNO Component that implements ``com.sun.star.script.XLibraryContainerExport`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XLibraryContainerExport``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XLibraryContainerExport
    def export_library(self, name: str, url: str, handler: XInteractionHandler) -> None:
        """
        Exports a library specified by Name to the location specified by the passed URL string.

        An interaction handler can be passed to be used for internal ucb operations.
        Exceptions not processed by this handler will be passed as com.sun.star.uno.Exception.
        If this parameter is null this applies to all exceptions thrown by ucb.

        If a library with the this name doesn't exist a com.sun.star.container.NoSuchElementException is thrown.

        Raises:
            com.sun.star.uno.Exception: ``Exception``
            com.sun.star.container.NoSuchElementException: ``NoSuchElementException``
        """
        self.__component.exportLibrary(name, url, handler)

    # endregion XLibraryContainerExport
