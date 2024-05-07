from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno
from com.sun.star.ucb import XSimpleFileAccess2

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.adapter.ucb.simple_file_access_partial import SimpleFileAccessPartial

if TYPE_CHECKING:
    from com.sun.star.io import XInputStream
    from ooodev.utils.type_var import UnoInterface


class SimpleFileAccess2Partial(SimpleFileAccessPartial):
    """
    Partial Class XSimpleFileAccess2.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XSimpleFileAccess2, interface: UnoInterface | None = XSimpleFileAccess2) -> None:
        """
        Constructor

        Args:
            component (XSimpleFileAccess2): UNO Component that implements ``com.sun.star.ucb.XSimpleFileAccess2`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XSimpleFileAccess2``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        SimpleFileAccessPartial.__init__(self, component=component, interface=None)
        self.__component = component

    # region XSimpleFileAccess2
    def write_file(self, file_url: str, data: XInputStream) -> None:
        """
        Overwrites the file content with the given data.

        If the file does not exist, it will be created.

        Raises:
            com.sun.star.uno.Exception: ``Exception``
        """
        return self.__component.writeFile(file_url, data)

    # endregion XSimpleFileAccess2
