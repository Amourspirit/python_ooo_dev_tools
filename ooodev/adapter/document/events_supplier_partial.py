from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.document import XEventsSupplier

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo


if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface
    from ooodev.adapter.container.name_replace_comp import NameReplaceComp


class EventsSupplierPartial:
    """
    Partial class for XEventsSupplier.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XEventsSupplier, interface: UnoInterface | None = XEventsSupplier) -> None:
        """
        Constructor

        Args:
            component (XEventsSupplier): UNO Component that implements ``com.sun.star.container.XEventsSupplier`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XEventsSupplier``.
        """

        self.__interface = interface
        self.__validate(component)
        self.__component = component

    def __validate(self, component: Any) -> None:
        """
        Validates the component.

        Args:
            component (Any): The component to be validated.
        """
        if self.__interface is None:
            return
        if not mLo.Lo.is_uno_interfaces(component, self.__interface):
            raise mEx.MissingInterfaceError(self.__interface)

    # region XEventsSupplier
    def get_events(self) -> NameReplaceComp:
        """
        Offers a list of event handlers which are be bound to events of this object
        """
        from ooodev.adapter.container.name_replace_comp import NameReplaceComp

        comp = self.__component.getEvents()
        return NameReplaceComp(component=comp)

    # endregion XEventsSupplier
