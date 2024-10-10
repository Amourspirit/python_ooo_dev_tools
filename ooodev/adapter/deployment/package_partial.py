from __future__ import annotations
from typing import Any, TYPE_CHECKING, Tuple
import uno

from com.sun.star.deployment import XPackage

from ooodev.adapter.util.modify_broadcaster_partial import ModifyBroadcasterPartial
from ooodev.adapter.lang.component_partial import ComponentPartial


if TYPE_CHECKING:
    from com.sun.star.beans import StringPair
    from com.sun.star.graphic import XGraphic
    from com.sun.star.deployment import XPackageTypeInfo
    from com.sun.star.task import XAbortChannel
    from com.sun.star.ucb import XCommandEnvironment
    from ooodev.utils.type_var import UnoInterface


class PackagePartial(ModifyBroadcasterPartial, ComponentPartial):
    """
    Partial class for XPackage.
    """

    def __init__(self, component: XPackage, interface: UnoInterface | None = XPackage) -> None:
        """
        Constructor

        Args:
            component (XPackage): UNO Component that implements ``com.sun.star.deployment.XPackage`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XPackage``.
        """

        ModifyBroadcasterPartial.__init__(self, component=component, interface=interface)
        ComponentPartial.__init__(self, component=component, interface=None)
        self.__component = component

    # region XPackage
    def check_dependencies(self, cmd_env: XCommandEnvironment) -> bool:
        """
        checks if the dependencies for this package are still satisfied

        After updating the OpenOffice.org, some dependencies for packages might no longer be satisfied.

        **since**

            OOo 3.2

        Raises:
            DeploymentException: ``DeploymentException``
            ExtensionRemovedException: ``ExtensionRemovedException``
            com.sun.star.ucb.CommandFailedException: ``CommandFailedException``
        """
        return self.__component.checkDependencies(cmd_env)

    def check_prerequisites(
        self, abort_channel: XAbortChannel, cmd_env: XCommandEnvironment, alreadyInstalled: bool
    ) -> int:
        """
        Checks if the package can be installed.

        Only if the return value is TRUE the package is allowed to be installed. In case of FALSE or in case of an exception,
        the package must be removed completely. After return of this function no code from the extension may be used anymore,
        so that the extension can be safely removed from the hard disk.

        Raises:
            DeploymentException: ``DeploymentException``
            ExtensionRemovedException: ``ExtensionRemovedException``
            com.sun.star.ucb.CommandFailedException: ``CommandFailedException``
            com.sun.star.ucb.CommandAbortedException: ``CommandAbortedException``
        """
        return self.__component.checkPrerequisites(abort_channel, cmd_env, alreadyInstalled)

    def create_abort_channel(self) -> XAbortChannel:
        """
        Creates a command channel to be used to asynchronously abort a command.
        """
        return self.__component.createAbortChannel()

    def export_to(
        self, dest_folder_url: str, newTitle: str, name_clash_action: int, cmd_env: XCommandEnvironment
    ) -> None:
        """
        Exports package to given destination URL.

        Raises:
            ExtensionRemovedException: ``ExtensionRemovedException``
            com.sun.star.ucb.CommandFailedException: ``CommandFailedException``
            com.sun.star.ucb.CommandAbortedException: ``CommandAbortedException``
            com.sun.star.ucb.ContentCreationException: ``ContentCreationException``
        """
        self.__component.exportTo(dest_folder_url, newTitle, name_clash_action, cmd_env)

    def get_bundle(self, abort_channel: XAbortChannel, cmd_env: XCommandEnvironment) -> Tuple[XPackage, ...]:
        """
        Gets packages of the bundle.

        If isRemoved() Returns TRUE then getBundle may return an empty sequence in case the object is not registered.

        Raises:
            DeploymentException: ``DeploymentException``
            com.sun.star.ucb.CommandFailedException: ``CommandFailedException``
            com.sun.star.ucb.CommandAbortedException: ``CommandAbortedException``
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
        """
        return self.__component.getBundle(abort_channel, cmd_env)

    def get_description(self) -> str:
        """
        Returns a description string to describe the package.

        Raises:
            ExtensionRemovedException: ``ExtensionRemovedException``
        """
        return self.__component.getDescription()

    def get_display_name(self) -> str:
        """
        Returns the display name of the package, e.g.

        for graphical user interfaces (GUI).

        Raises:
            ExtensionRemovedException: ``ExtensionRemovedException``
        """
        return self.__component.getDisplayName()

    def get_icon(self, high_contrast: bool) -> XGraphic:
        """
        Returns an icon for a package.

        Raises:
            ExtensionRemovedException: ``ExtensionRemovedException``
        """
        return self.__component.getIcon(high_contrast)

    def get_identifier(self) -> Any:
        """
        Returns the unique extension identifier.
        """
        return self.__component.getIdentifier()

    def get_license_text(self) -> str:
        """
        Returns a string containing the license text.

        Raises:
            DeploymentException: ``DeploymentException``
            ExtensionRemovedException: ``ExtensionRemovedException``
        """
        return self.__component.getLicenseText()

    def get_name(self) -> str:
        """
        Returns the file name of the package.
        """
        return self.__component.getName()

    def get_package_type(self) -> XPackageTypeInfo:
        """
        Returns the XPackageTypeInfo, e.g.

        media-type etc.
        """
        return self.__component.getPackageType()

    def get_publisher_info(self) -> StringPair:
        """
        Returns the publisher info for the package, the strings might be empty, if there is no publisher

        com.sun.star.beans.StringPair.First represents the publisher name and com.sun.star.beans.StringPair.Second represents the URL to the publisher.

        Raises:
            ExtensionRemovedException: ``ExtensionRemovedException``
        """
        return self.__component.getPublisherInfo()

    def get_registration_data_url(self) -> Any:
        """
        return a URL to a directory which contains the registration data.

        This data may be created when calling XPackage.registerPackage(). If this is the case is indicated by com.sun.star.beans.Optional.IsPresent of the return value. If registration data are created during registration, but the package is currently not registered, for example after calling XPackage.revokePackage(), then com.sun.star.beans.Optional.IsPresent is TRUE and the com.sun.star.beans.Optional.Value may be an empty string.

        Raises:
            DeploymentException: ``DeploymentException``
            ExtensionRemovedException: ``ExtensionRemovedException``
        """
        return self.__component.getRegistrationDataURL()

    def get_repository_name(self) -> str:
        """
        Returns the name of the repository where this object comes from.
        """
        return self.__component.getRepositoryName()

    def get_url(self) -> str:
        """
        Returns the location of the package.
        """
        return self.__component.getURL()

    def get_update_information_ur_ls(self) -> Tuple[str, ...]:
        """
        Returns a sequence of update information URLs.

        The sequence may be empty in case no update information is available. If the sequence contains more than one URL, the extra URLs must mirror the information available at the first URL.

        Raises:
            ExtensionRemovedException: ``ExtensionRemovedException``
        """
        return self.__component.getUpdateInformationURLs()

    def get_version(self) -> str:
        """
        Returns the textual version representation of the package.

        A textual version representation is a finite string following the ``BNFversion .= [element (. element)*]element .= (0 | 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9)+``

        Raises:
            ExtensionRemovedException: ``ExtensionRemovedException``
        """
        return self.__component.getVersion()

    def is_bundle(self) -> bool:
        """
        Reflects whether this package is a bundle of one or more packages, e.g.

        a zip (legacy) package file or a document hosting script packages.
        """
        return self.__component.isBundle()

    def is_registered(self, abort_channel: XAbortChannel, cmd_env: XCommandEnvironment) -> Any:
        """
        Determines whether the package is currently registered, i.e.

        whether it is active.

        Raises:
            DeploymentException: ``DeploymentException``
            com.sun.star.ucb.CommandFailedException: ``CommandFailedException``
            com.sun.star.ucb.CommandAbortedException: ``CommandAbortedException``
        """
        return self.__component.isRegistered(abort_channel, cmd_env)

    def is_removed(self) -> bool:
        """
        Indicates if this object represents a removed extension or extension item.

        This is the case when it was created by providing ``True`` for the removed parameter in the function ``XPackageRegistry.bindPackage()``.
        """
        return self.__component.isRemoved()

    def register_package(self, startup: bool, abort_channel: XAbortChannel, cmd_env: XCommandEnvironment) -> None:
        """
        Registers this XPackage.

        NEVER call this directly. This is done by the extension manager if necessary.

        Raises:
            DeploymentException: ``DeploymentException``
            ExtensionRemovedException: ``ExtensionRemovedException``
            com.sun.star.ucb.CommandFailedException: ``CommandFailedException``
            com.sun.star.ucb.CommandAbortedException: ``CommandAbortedException``
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
        """
        return self.__component.registerPackage(startup, abort_channel, cmd_env)

    def revoke_package(self, startup: bool, abort_channel: XAbortChannel, cmd_env: XCommandEnvironment) -> None:
        """
        revokes this XPackage.

        NEVER call this directly. This is done by the extension manager if necessary.

        Raises:
            DeploymentException: ``DeploymentException``
            com.sun.star.ucb.CommandFailedException: ``CommandFailedException``
            com.sun.star.ucb.CommandAbortedException: ``CommandAbortedException``
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
        """
        self.__component.revokePackage(startup, abort_channel, cmd_env)

    # endregion XPackage
