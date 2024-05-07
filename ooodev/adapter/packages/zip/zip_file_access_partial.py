from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno
from com.sun.star.packages.zip import XZipFileAccess

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.io import XInputStream
    from ooodev.utils.type_var import UnoInterface


class ZipFileAccessPartial:
    """
    Partial class for XZipFileAccess.
    """

    def __init__(self, component: XZipFileAccess, interface: UnoInterface | None = XZipFileAccess) -> None:
        """
        Constructor

        Args:
            component (XZipFileAccess): UNO Component that implements ``com.sun.star.container.XZipFileAccess`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XZipFileAccess``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XZipFileAccess
    def get_stream_by_pattern(self, pattern: str) -> XInputStream:
        """
        Allows to get stream by specifying a pattern.

        The first stream with a name that fits to the pattern will be returned.
        The pattern allows to use ``*`` wildcard symbol.
        If the name contains ``*`` or ``\\`` symbols itself they must guarded with backslash ``\\``.
        The slashes have no special meaning here so they can be replaced by wildcards also.

        Raises:
            com.sun.star.container.NoSuchElementException: ``NoSuchElementException``
            com.sun.star.io.IOException: ``IOException``
            com.sun.star.packages.WrongPasswordException: ``WrongPasswordException``
            com.sun.star.packages.zip.ZipException: ``ZipException``
        """
        return self.__component.getStreamByPattern(pattern)

    # endregion XZipFileAccess
