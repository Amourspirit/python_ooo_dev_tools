from __future__ import annotations
from typing import Any, TYPE_CHECKING, Tuple
import uno

from com.sun.star.deployment import XPackageManager

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.deployment import XPackageTypeInfo
    from com.sun.star.beans import NamedValue
    from com.sun.star.deployment import XPackage
    from com.sun.star.task import XAbortChannel
    from com.sun.star.ucb import XCommandEnvironment
    from ooodev.utils.type_var import UnoInterface


class PackageManagerPartial:
    """
    Partial class for XPackageManager.
    """

    def __init__(self, component: XPackageManager, interface: UnoInterface | None = XPackageManager) -> None:
        """
        Constructor

        Args:
            component (XPackageManager): UNO Component that implements ``com.sun.star.deployment.XPackageManager`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XPackageManager``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XPackageManager
    def add_package(
        self,
        url: str,
        media_type: str,
        abort_channel: XAbortChannel,
        cmd_env: XCommandEnvironment,
        *properties: NamedValue,
    ) -> XPackage:
        """
        adds a UNO package.

        The properties argument is currently only used to suppress the license information for shared extensions.

        Raises:
            DeploymentException: ``DeploymentException``
            com.sun.star.ucb.CommandFailedException: ``CommandFailedException``
            com.sun.star.ucb.CommandAbortedException: ``CommandAbortedException``
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
        """
        return self.__component.addPackage(url, properties, media_type, abort_channel, cmd_env)

    def check_prerequisites(
        self,
        extension: XPackage,
        abort_channel: XAbortChannel,
        cmd_env: XCommandEnvironment,
    ) -> int:
        """
        checks if the extension can be used.

        The extension must be managed by this package manager, that is, it must be recorded in its database. The package manager calls XPackage.checkPrerequisites and updates its data base with the result. The result, which is from Prerequisites will be returned.

        Raises:
            DeploymentException: ``DeploymentException``
            com.sun.star.ucb.CommandFailedException: ``CommandFailedException``
            com.sun.star.ucb.CommandAbortedException: ``CommandAbortedException``
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
        """
        return self.__component.checkPrerequisites(extension, abort_channel, cmd_env)

    def create_abort_channel(self) -> XAbortChannel:
        """
        creates a command channel to be used to asynchronously abort a command.
        """
        return self.__component.createAbortChannel()

    def get_context(self) -> str:
        """
        returns the underlying deployment context, that is, the name of the repository.
        """
        return self.__component.getContext()

    def get_deployed_package(self, identifier: str, file_name: str, cmd_env: XCommandEnvironment) -> XPackage:
        """
        gets a deployed package.

        Raises:
            DeploymentException: ``DeploymentException``
            com.sun.star.ucb.CommandFailedException: ``CommandFailedException``
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
        """
        return self.__component.getDeployedPackage(identifier, file_name, cmd_env)

    def get_deployed_packages(
        self, abort_channel: XAbortChannel, cmd_env: XCommandEnvironment
    ) -> Tuple[XPackage, ...]:
        """
        gets all currently deployed packages.

        Raises:
            DeploymentException: ``DeploymentException``
            com.sun.star.ucb.CommandFailedException: ``CommandFailedException``
            com.sun.star.ucb.CommandAbortedException: ``CommandAbortedException``
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
        """
        return self.__component.getDeployedPackages(abort_channel, cmd_env)

    def get_extensions_with_unaccepted_licenses(self, cmd_env: XCommandEnvironment) -> Tuple[XPackage, ...]:
        """
        returns all extensions which are currently not in use because the user did not accept the license.

        The function will not return any object for the user repository, because a user extension will not be kept in the user repository if its license is declined. Only extensions which are registered at start-up of OOo, that is, shared and bundled extensions, can be returned.

        Extensions which allow the license to be suppressed, that is, it does not need to be displayed, and which are installed with the corresponding option, are also not returned.

        Raises:
            DeploymentException: ``DeploymentException``
        """
        return self.__component.getExtensionsWithUnacceptedLicenses(cmd_env)

    def get_supported_package_types(self) -> Tuple[XPackageTypeInfo, ...]:
        """
        gets the supported XPackageTypeInfos.
        """
        return self.__component.getSupportedPackageTypes()

    def import_extension(
        self,
        extension: XPackage,
        abort_channel: XAbortChannel,
        cmd_env: XCommandEnvironment,
    ) -> XPackage:
        """
        adds an extension.

        This copies the extension. If it was from the same repository, which is represented by this XPackageManager interface, then nothing happens.

        Raises:
            DeploymentException: ``DeploymentException``
            com.sun.star.ucb.CommandFailedException: ``CommandFailedException``
            com.sun.star.ucb.CommandAbortedException: ``CommandAbortedException``
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
        """
        return self.__component.importExtension(extension, abort_channel, cmd_env)

    def is_read_only(self) -> bool:
        """
        indicates that this implementation cannot be used for tasks which require write access to the location where the extensions are installed.

        Normally one would call a method and handle the exception if writing failed. However, a GUI interface may need to know beforehand if writing is allowed. For example, the Extension Manager dialog needs to enable / disable the Add button depending if the user has write permission. Only the XPackageManager implementation knows the location of the installed extensions. Therefore it is not possible to check \"externally\" for write permission.
        """
        return self.__component.isReadOnly()

    def reinstall_deployed_packages(
        self, force: bool, abort_channel: XAbortChannel, cmd_env: XCommandEnvironment
    ) -> None:
        """
        Expert feature: erases the underlying registry cache and reinstalls all previously added packages.

        Please keep in mind that all registration status get lost.

        Please use this in case of suspected cache inconsistencies only.

        Raises:
            DeploymentException: ``DeploymentException``
            com.sun.star.ucb.CommandFailedException: ``CommandFailedException``
            com.sun.star.ucb.CommandAbortedException: ``CommandAbortedException``
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
        """
        return self.__component.reinstallDeployedPackages(force, abort_channel, cmd_env)

    def remove_package(
        self,
        identifier: str,
        fileName: str,
        abort_channel: XAbortChannel,
        cmd_env: XCommandEnvironment,
    ) -> None:
        """
        removes a UNO package.

        Raises:
            DeploymentException: ``DeploymentException``
            com.sun.star.ucb.CommandFailedException: ``CommandFailedException``
            com.sun.star.ucb.CommandAbortedException: ``CommandAbortedException``
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
        """
        self.__component.removePackage(identifier, fileName, abort_channel, cmd_env)

    def synchronize(self, abort_channel: XAbortChannel, cmd_env: XCommandEnvironment) -> bool:
        """
        synchronizes the extension database with the contents of the extensions folder.

        Added extensions will be added to the database and removed extensions will be removed from the database.

        Raises:
            DeploymentException: ``DeploymentException``
            com.sun.star.ucb.ContentCreationException: ``ContentCreationException``
            com.sun.star.ucb.CommandFailedException: ``CommandFailedException``
            com.sun.star.ucb.CommandAbortedException: ``CommandAbortedException``
        """
        return self.__component.synchronize(abort_channel, cmd_env)

    # endregion XPackageManager
