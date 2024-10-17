from __future__ import annotations
from typing import Any, TYPE_CHECKING, Tuple

from com.sun.star.deployment import XUpdateInformationProvider

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.container import XEnumeration
    from com.sun.star.task import XInteractionHandler
    from com.sun.star.xml.dom import XElement
    from ooodev.utils.type_var import UnoInterface


class UpdateInformationProviderPartial:
    """
    Partial class for XUpdateInformationProvider.
    """

    def __init__(
        self, component: XUpdateInformationProvider, interface: UnoInterface | None = XUpdateInformationProvider
    ) -> None:
        """
        Constructor

        Args:
            component (XUpdateInformationProvider): UNO Component that implements ``com.sun.star.deployment.XUpdateInformationProvider`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XUpdateInformationProvider``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XUpdateInformationProvider
    def cancel(self) -> None:
        """
        interrupts a ``get_update_information`` call and let's it return immediately.
        """
        return self.__component.cancel()

    def get_update_information(self, repositories: Tuple[str, ...], extension_id: str) -> Tuple[XElement, ...]:
        """
        get update information for a specific extension or all available information from a repository.

        Raises:
            com.sun.star.uno.Exception: ``Exception``
        """
        return self.__component.getUpdateInformation(repositories, extension_id)

    def get_update_information_enumeration(self, repositories: Tuple[str, ...], extension_id: str) -> XEnumeration:
        """
        get update information for a specific extension or all available information from a repository.

        Raises:
            com.sun.star.uno.Exception: ``Exception``
        """
        return self.__component.getUpdateInformationEnumeration(repositories, extension_id)

    def set_interaction_handler(self, handler: XInteractionHandler) -> None:
        """
        Sets an interaction handler to be used for further operations.

        A default interaction handler is available as service com.sun.star.task.InteractionHandler. The documentation of this service also contains further information about the interaction handler concept.
        """
        self.__component.setInteractionHandler(handler)

    # endregion XUpdateInformationProvider
