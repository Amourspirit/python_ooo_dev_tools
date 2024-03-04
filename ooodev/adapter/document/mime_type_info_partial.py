from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.document import XMimeTypeInfo

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class MimeTypeInfoPartial:
    """
    Partial class for XMimeTypeInfo.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XMimeTypeInfo, interface: UnoInterface | None = XMimeTypeInfo) -> None:
        """
        Constructor

        Args:
            component (XMimeTypeInfo): UNO Component that implements ``com.sun.star.container.XMimeTypeInfo`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XMimeTypeInfo``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XMimeTypeInfo
    def get_supported_mime_type_names(self) -> tuple[str, ...]:
        """
        Gets the supported MIME type names.

        Returns:
            tuple[str, ...]: The supported MIME type names.
        """
        return self.__component.getSupportedMimeTypeNames()

    def supports_mime_type(self, mime_type: str) -> bool:
        """
        Checks whether the specified MIME type is supported.

        Args:
            mime_type (str): The MIME type to be checked.

        Returns:
            bool: ``True`` if the specified MIME type is supported, otherwise ``False``.
        """
        return self.__component.supportsMimeType(mime_type)

    # endregion XMimeTypeInfo
