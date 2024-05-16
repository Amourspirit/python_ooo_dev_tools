from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno
from com.sun.star.script import XLibraryQueryExecutable

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class LibraryQueryExecutablePartial:
    """
    Partial class for XLibraryQueryExecutable.
    """

    def __init__(
        self, component: XLibraryQueryExecutable, interface: UnoInterface | None = XLibraryQueryExecutable
    ) -> None:
        """
        Constructor

        Args:
            component (XLibraryQueryExecutable): UNO Component that implements ``com.sun.star.script.XLibraryQueryExecutable`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XLibraryQueryExecutable``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XLibraryQueryExecutable
    def has_executable_code(self, name: str) -> bool:
        """
        Gets whether the library contains executable code for the specified name.
        """
        return self.__component.HasExecutableCode(name)

    # endregion XLibraryQueryExecutable
