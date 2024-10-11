from __future__ import annotations
from typing import Any, TYPE_CHECKING, Tuple
import uno

from com.sun.star.deployment import XPackageInformationProvider

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class PackageInformationProviderPartial:
    """
    Partial class for XPackageInformationProvider.
    """

    def __init__(
        self,
        component: XPackageInformationProvider,
        interface: UnoInterface | None = XPackageInformationProvider,
    ) -> None:
        """
        Constructor

        Args:
            component (XPackageInformationProvider): UNO Component that implements ``com.sun.star.deployment.XPackageInformationProvider`` interface.
            interface (UnoInterface, optional): Interface to be validated. Defaults to ``XPackageInformationProvider``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XPackageInformationProvider

    def get_extension_list(self) -> Tuple[Tuple[str, ...], ...]:
        """
        Returns a list of all installed extensions with their versions.
        """
        return self.__component.getExtensionList()

    def get_package_location(self, extension_id: str) -> str:
        """
        Gets package information for a specific extension.
        """
        return self.__component.getPackageLocation(extension_id)

    def is_update_available(self, extension_id: str) -> Tuple[Tuple[str, ...], ...]:
        """
        Checks if there are updates available for an extension.
        """
        return self.__component.isUpdateAvailable(extension_id)

    # endregion XPackageInformationProvider
