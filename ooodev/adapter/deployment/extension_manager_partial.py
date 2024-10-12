from __future__ import annotations
from typing import TYPE_CHECKING, Tuple
import uno

from com.sun.star.deployment import XExtensionManager

from ooodev.adapter.util.modify_broadcaster_partial import ModifyBroadcasterPartial
from ooodev.adapter.lang.component_partial import ComponentPartial


if TYPE_CHECKING:
    from com.sun.star.deployment import XPackageTypeInfo
    from com.sun.star.beans import NamedValue
    from com.sun.star.deployment import XPackage
    from com.sun.star.task import XAbortChannel
    from com.sun.star.ucb import XCommandEnvironment
    from ooodev.utils.type_var import UnoInterface


class ExtensionManagerPartial(ModifyBroadcasterPartial, ComponentPartial):
    """
    Partial class for XExtensionManager.
    """

    def __init__(self, component: XExtensionManager, interface: UnoInterface | None = XExtensionManager) -> None:
        """
        Constructor

        Args:
            component (XExtensionManager): UNO Component that implements ``com.sun.star.deployment.XExtensionManager`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XExtensionManager``.
        """

        ModifyBroadcasterPartial.__init__(self, component=component, interface=interface)
        ComponentPartial.__init__(self, component=component, interface=None)
        self.__component = component

    # region XExtensionManager
    def add_extension(
        self,
        url: str,
        repository: str,
        abort_channel: XAbortChannel,
        cmd_env: XCommandEnvironment,
        *properties: NamedValue,
    ) -> XPackage:
        """
        adds an extension.

        The properties argument is currently only used to suppress the license information for shared extensions.

        Raises:
            DeploymentException: ``DeploymentException``
            com.sun.star.ucb.CommandFailedException: ``CommandFailedException``
            com.sun.star.ucb.CommandAbortedException: ``CommandAbortedException``
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
        """
        return self.__component.addExtension(url, properties, repository, abort_channel, cmd_env)

    def check_prerequisites_and_enable(
        self,
        extension: XPackage,
        abort_channel: XAbortChannel,
        cmd_env: XCommandEnvironment,
    ) -> int:
        """
        check if all prerequisites for the extension are fulfilled and activates it, if possible.

        Raises:
            DeploymentException: ``DeploymentException``
            com.sun.star.ucb.CommandFailedException: ``CommandFailedException``
            com.sun.star.ucb.CommandAbortedException: ``CommandAbortedException``
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
        """
        return self.__component.checkPrerequisitesAndEnable(extension, abort_channel, cmd_env)

    def create_abort_channel(self) -> XAbortChannel:
        """
        creates a command channel to be used to asynchronously abort a command.
        """
        return self.__component.createAbortChannel()

    def disable_extension(
        self,
        extension: XPackage,
        abort_channel: XAbortChannel,
        cmd_env: XCommandEnvironment,
    ) -> None:
        """
        disable an extension.

        If the extension is not from the user repository then an IllegalArgumentException is thrown.

        Raises:
            DeploymentException: ``DeploymentException``
            com.sun.star.ucb.CommandFailedException: ``CommandFailedException``
            com.sun.star.ucb.CommandAbortedException: ``CommandAbortedException``
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
        """
        return self.__component.disableExtension(extension, abort_channel, cmd_env)

    def enable_extension(
        self,
        extension: XPackage,
        abort_channel: XAbortChannel,
        cmd_env: XCommandEnvironment,
    ) -> None:
        """
        enable an extension.

        If the extension is not from the user repository then an IllegalArgumentException is thrown.

        Raises:
            DeploymentException: ``DeploymentException``
            com.sun.star.ucb.CommandFailedException: ``CommandFailedException``
            com.sun.star.ucb.CommandAbortedException: ``CommandAbortedException``
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
        """
        return self.__component.enableExtension(extension, abort_channel, cmd_env)

    def get_all_extensions(
        self, abort_channel: XAbortChannel, cmd_env: XCommandEnvironment
    ) -> Tuple[Tuple[XPackage, ...], ...]:
        """
        returns a sequence containing all installed extensions.

        The members of the returned sequence correspond to an extension with a particular extension identifier. The members are also sequences which contain as many elements as there are repositories. Those are ordered according to the priority of the repository. That is, the first member is the extension from the user repository, the second is from the shared repository and the last is from the bundled repository.

        Raises:
            DeploymentException: ``DeploymentException``
            com.sun.star.ucb.CommandFailedException: ``CommandFailedException``
            com.sun.star.ucb.CommandAbortedException: ``CommandAbortedException``
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
        """
        return self.__component.getAllExtensions(abort_channel, cmd_env)

    def get_deployed_extension(
        self, repository: str, identifier: str, file_name: str, cmd_env: XCommandEnvironment
    ) -> XPackage:
        """
        gets an installed extensions.

        Raises:
            DeploymentException: ``DeploymentException``
            com.sun.star.ucb.CommandFailedException: ``CommandFailedException``
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
        """
        return self.__component.getDeployedExtension(repository, identifier, file_name, cmd_env)

    def get_deployed_extensions(
        self, repository: str, abort_channel: XAbortChannel, cmd_env: XCommandEnvironment
    ) -> Tuple[XPackage, ...]:
        """
        gets all currently installed extensions, including disabled user extensions.

        Raises:
            DeploymentException: ``DeploymentException``
            com.sun.star.ucb.CommandFailedException: ``CommandFailedException``
            com.sun.star.ucb.CommandAbortedException: ``CommandAbortedException``
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
        """
        return self.__component.getDeployedExtensions(repository, abort_channel, cmd_env)

    def get_extensions_with_same_identifier(
        self, identifier: str, file_name: str, cmd_env: XCommandEnvironment
    ) -> Tuple[XPackage, ...]:
        """
        gets all extensions with the same identifier from all repositories.

        The extension at the first position in the returned sequence represents the extension from the user repository. The next element is from the shared and the last one is from the bundled repository. If one repository does not contain this extension, then the respective element is a null reference.

        Raises:
            DeploymentException: ``DeploymentException``
            com.sun.star.ucb.CommandFailedException: ``CommandFailedException``
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
        """
        return self.__component.getExtensionsWithSameIdentifier(identifier, file_name, cmd_env)

    def get_extensions_with_unaccepted_licenses(
        self, repository: str, cmd_env: XCommandEnvironment
    ) -> Tuple[XPackage, ...]:
        """
        returns all extensions which are currently not in use because the user did not accept the license.

        The function will not return any object for the user repository, because a user extension will not be kept in the user repository if its license is declined. Only extensions which are registered at start-up of OOo, that is, shared and bundled extensions, can be returned.

        Extensions which allow the license to be suppressed, that is, it does not need to be displayed, and which are installed with the corresponding option, are also not returned.

        Extensions returned by these functions are not returned by XExtensionManager.getDeployedExtension() XExtensionManager.getDeployedExtensions() XExtensionManager.getAllExtensions() XExtensionManager.getExtensionsWithSameIdentifier()

        Raises:
            DeploymentException: ``DeploymentException``
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
        """
        return self.__component.getExtensionsWithUnacceptedLicenses(repository, cmd_env)

    def get_supported_package_types(self) -> Tuple[XPackageTypeInfo, ...]:
        """
        gets the supported XPackageTypeInfos.
        """
        return self.__component.getSupportedPackageTypes()

    def is_read_only_repository(self, repository: str) -> bool:
        """
        determines if the current user has write access to the extensions folder of the repository.
        """
        return self.__component.isReadOnlyRepository(repository)

    def reinstall_deployed_extensions(
        self,
        force: bool,
        repository: str,
        abort_channel: XAbortChannel,
        cmd_env: XCommandEnvironment,
    ) -> None:
        """
        Expert feature: erases the underlying registry cache and reinstalls all previously added extensions.

        Please keep in mind that all registration status get lost.

        Please use this in case of suspected cache inconsistencies only.

        Raises:
            DeploymentException: ``DeploymentException``
            com.sun.star.ucb.CommandFailedException: ``CommandFailedException``
            com.sun.star.ucb.CommandAbortedException: ``CommandAbortedException``
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
        """
        return self.__component.reinstallDeployedExtensions(force, repository, abort_channel, cmd_env)

    def remove_extension(
        self,
        identifier: str,
        file_name: str,
        repository: str,
        abort_channel: XAbortChannel,
        cmd_env: XCommandEnvironment,
    ) -> None:
        """
        removes an extension.

        Raises:
            DeploymentException: ``DeploymentException``
            com.sun.star.ucb.CommandFailedException: ``CommandFailedException``
            com.sun.star.ucb.CommandAbortedException: ``CommandAbortedException``
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
        """
        self.__component.removeExtension(identifier, file_name, repository, abort_channel, cmd_env)

    def synchronize(self, abort_channel: XAbortChannel, cmd_env: XCommandEnvironment) -> bool:
        """
        synchronizes the extension database with the contents of the extensions folder of shared and bundled extensions.

        Added extensions will be added to the database and removed extensions will be removed from the database. The active extensions are determined. That is, shared or bundled extensions are not necessarily registered (XPackage.registerPackage()).

        Raises:
            DeploymentException: ``DeploymentException``
            com.sun.star.ucb.CommandFailedException: ``CommandFailedException``
            com.sun.star.ucb.CommandAbortedException: ``CommandAbortedException``
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
        """
        return self.__component.synchronize(abort_channel, cmd_env)

    # endregion XExtensionManager
