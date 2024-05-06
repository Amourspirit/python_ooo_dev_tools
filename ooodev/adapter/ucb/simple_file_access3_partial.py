from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno
from com.sun.star.ucb import XSimpleFileAccess3

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.adapter.ucb.simple_file_access2_partial import SimpleFileAccess2Partial

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class SimpleFileAccess3Partial(SimpleFileAccess2Partial):
    """
    Partial Class XSimpleFileAccess3.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XSimpleFileAccess3, interface: UnoInterface | None = XSimpleFileAccess3) -> None:
        """
        Constructor

        Args:
            component (XSimpleFileAccess3): UNO Component that implements ``com.sun.star.ucb.XSimpleFileAccess3`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XSimpleFileAccess3``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        SimpleFileAccess2Partial.__init__(self, component=component, interface=None)
        self.__component = component

    # region XSimpleFileAccess3
    def is_hidden(self, file_url: str) -> bool:
        """
        Checks if a file is ``hidden``.

        Raises:
            com.sun.star.ucb.CommandAbortedException: ``CommandAbortedException``
            com.sun.star.uno.Exception: ``Exception``
        """
        return self.__component.isHidden(file_url)

    def set_hidden(self, file_url: str, hidden: bool) -> None:
        """
        Sets the ``hidden`` of a file according to the boolean parameter, if the actual process has the right to do so and the used operation system supports this operation.

        Raises:
            com.sun.star.ucb.CommandAbortedException: ``CommandAbortedException``
            com.sun.star.uno.Exception: ``Exception``
        """
        self.__component.setHidden(file_url, hidden)

    # endregion XSimpleFileAccess3
