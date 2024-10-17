from __future__ import annotations
from typing import Any, TYPE_CHECKING

from com.sun.star.deployment import XPackageManagerFactory

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.deployment import XPackageManager
    from ooodev.utils.type_var import UnoInterface


class PackageManagerFactoryPartial:
    """
    Partial class for XPackageManagerFactory.
    """

    def __init__(
        self, component: XPackageManagerFactory, interface: UnoInterface | None = XPackageManagerFactory
    ) -> None:
        """
        Constructor

        Args:
            component (XPackageManagerFactory): UNO Component that implements ``com.sun.star.deployment.XPackageManagerFactory`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XPackageManagerFactory``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XPackageManagerFactory
    def get_package_manager(self, context: str) -> XPackageManager:
        """
        Method to create (or reusing and already existing) XPackageManager object to add or remove UNO packages persistently.

        Packages for context strings ``user`` and ``shared`` will be registered and revoked persistently.

        Context strings other than ``user``, ``shared`` will last in an ``com.sun.star.lang.IllegalArgumentException``.

        Raises:
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
        """
        return self.__component.getPackageManager(context)

    # endregion XPackageManagerFactory
