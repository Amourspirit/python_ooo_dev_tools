from __future__ import annotations
from typing import Any, TYPE_CHECKING, Type
import uno
from com.sun.star.packages.zip import XZipFileAccess2
from ooodev.adapter.io.input_stream_comp import InputStreamComp

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.adapter.packages.zip.zip_file_access_partial import ZipFileAccessPartial
from ooodev.adapter.container.name_access_partial import NameAccessPartial

if TYPE_CHECKING:
    from com.sun.star.io import XInputStream
    from ooodev.utils.type_var import UnoInterface


class ZipFileAccess2Partial(ZipFileAccessPartial, NameAccessPartial[InputStreamComp]):
    """
    Partial class for XZipFileAccess2.
    """

    def __init__(self, component: XZipFileAccess2, interface: UnoInterface | None = XZipFileAccess2) -> None:
        """
        Constructor

        Args:
            component (XZipFileAccess2): UNO Component that implements ``com.sun.star.container.XZipFileAccess2`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XZipFileAccess2``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        ZipFileAccessPartial.__init__(self, component=component, interface=None)
        NameAccessPartial.__init__(self, component=component, interface=None)

    # region XNameAccess

    def get_by_name(self, name: str) -> InputStreamComp:
        """
        Gets the element with the specified name.

        Args:
            name (str): The name of the element.

        Returns:
            Any: The element with the specified name.
        """
        result = super().get_by_name(name)
        if result is None:
            return None  # type: ignore
        return InputStreamComp(result)  # type: ignore

    # endregion XNameAccess

    # region XElementAccess
    def get_element_type(self) -> Type[InputStreamComp]:
        """
        Gets the type of the elements contained in the container.

        Returns:
            Any: The type of the elements. ``None``  means that it is a multi-type container and you cannot determine the exact types with this interface.
        """
        return InputStreamComp

    # endregion XElementAccess
