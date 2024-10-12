from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.deployment import XPackageTypeInfo

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class PackageTypeInfoPartial:
    """
    Partial class for XPackageTypeInfo.
    """

    def __init__(
        self,
        component: XPackageTypeInfo,
        interface: UnoInterface | None = XPackageTypeInfo,
    ) -> None:
        """
        Constructor.

        Args:
            component (XPackageTypeInfo): UNO component implementing the `com.sun.star.deployment.XPackageTypeInfo` interface.
            interface (UnoInterface, optional): Interface to validate against. Defaults to `XPackageTypeInfo`.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XPackageTypeInfo

    def get_description(self) -> str:
        """
        Returns a description string to describe a package type.

        Raises:
            ExtensionRemovedException: If the extension has been removed.
        """
        return self.__component.getDescription()

    def get_file_filter(self) -> str:
        """
        Returns a file filter string for the file picker user interface.

        Both the short description string and file filter string will be passed to
        `com.sun.star.ui.dialogs.XFilterManager.appendFilter()`.
        """
        return self.__component.getFileFilter()

    def get_icon(self, high_contrast: bool, small_icon: bool) -> Any:
        """
        Returns an icon for a package.

        Args:
            high_contrast (bool): Whether to use a high-contrast icon.
            small_icon (bool): Whether to use a small icon.

        Returns:
            Any: The icon representation.
        """
        return self.__component.getIcon(high_contrast, small_icon)

    def get_media_type(self) -> str:
        """
        Returns the media type of a package, e.g., 'application/vnd.sun.star.basic-script'.

        Returns:
            str: The media type of the package.
        """
        return self.__component.getMediaType()

    def get_short_description(self) -> str:
        """
        Returns a short description string to describe a package type (one line only).

        Returns:
            str: The short description of the package type.
        """
        return self.__component.getShortDescription()

    # endregion XPackageTypeInfo
