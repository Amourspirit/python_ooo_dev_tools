from __future__ import annotations
from typing import Any, TYPE_CHECKING, Tuple

from com.sun.star.deployment import XPackageRegistry

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.deployment import XPackage
    from com.sun.star.deployment import XPackageTypeInfo
    from com.sun.star.ucb import XCommandEnvironment
    from ooodev.utils.type_var import UnoInterface


class PackageRegistryPartial:
    """
    Partial class for XPackageRegistry.
    """

    def __init__(self, component: XPackageRegistry, interface: UnoInterface | None = XPackageRegistry) -> None:
        """
        Constructor

        Args:
            component (XPackageRegistry): UNO Component that implements ``com.sun.star.deployment.XPackageRegistry`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XPackageRegistry``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XPackageRegistry
    def bind_package(
        self, url: str, media_type: str, removed: bool, identifier: str, cmd_env: XCommandEnvironment
    ) -> XPackage:
        """
        binds a package URL to a XPackage handle.

        The returned UNO package handle ought to late-initialize itself, thus the process of binding must not be an expensive operation, because it is not abortable.

        Calling the function several time with the same parameters must result in returning the same object.

        The file or folder at the location where url points to may not exist or it was replaced. This can happen, for example, when a bundled extension was removed by the setup and a user later starts OOo. Then the user data may still contain all registration data of that extension, but the actual extension files do not exist anymore. The registration data must then be cleaned of all the remains of that extension. To do that one creates an XPackage object on behalf of that extension and calls XPackage.revokePackage(). The parameter removed indicates this case. The returned object may not rely on the file or folder to which refers url. Instead it must use previously saved data to successfully carry out the revocation of this object (XPackage.revokePackage()).

        The implementation must ensure that there is only one instance of XPackage for the same url at any time. Therefore calling bindPackage() again with the same url but different mediaType (the exception is, if previously an empty string was provided to cause the determination of the media type) or removed parameters will cause an exception. A com.sun.star.lang.IllegalArgumentException will be thrown in case of a different mediaType parameter and a InvalidRemovedParameterException is thrown if the removed parameter is different.

        The identifier parameter must be provided when removed = true. If not, then an com.sun.star.lang.IllegalArgumentException will be thrown.

        Raises:
            DeploymentException: ``DeploymentException``
            InvalidRemovedParameterException: ``InvalidRemovedParameterException``
            com.sun.star.ucb.CommandFailedException: ``CommandFailedException``
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
        """
        return self.__component.bindPackage(url, media_type, removed, identifier, cmd_env)

    def get_supported_package_types(self) -> Tuple[XPackageTypeInfo, ...]:
        """
        gets the supported XPackageTypeInfos.
        """
        return self.__component.getSupportedPackageTypes()

    def packageRemoved(self, url: str, media_type: str) -> None:
        """

        Raises:
            DeploymentException: ``DeploymentException``
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
        """
        return self.__component.packageRemoved(url, media_type)

    # endregion XPackageRegistry
