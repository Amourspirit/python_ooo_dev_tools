from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.document import XEmbeddedObjectSupplier

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo


if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface
    from com.sun.star.lang import XComponent


class EmbeddedObjectSupplierPartial:
    """
    Partial class for XEmbeddedObjectSupplier.
    """

    # pylint: disable=unused-argument

    def __init__(
        self, component: XEmbeddedObjectSupplier, interface: UnoInterface | None = XEmbeddedObjectSupplier
    ) -> None:
        """
        Constructor

        Args:
            component (XEmbeddedObjectSupplier): UNO Component that implements ``com.sun.star.container.XEmbeddedObjectSupplier`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XEmbeddedObjectSupplier``.
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

    # region XEmbeddedObjectSupplier
    def get_embedded_object(self) -> XComponent:
        """
        returns the object which is embedded into this object.
        """
        return self.__component.getEmbeddedObject()

    # endregion XEmbeddedObjectSupplier
