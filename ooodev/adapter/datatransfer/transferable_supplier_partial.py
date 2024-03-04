from __future__ import annotations
from typing import Any, TYPE_CHECKING

import uno
from com.sun.star.datatransfer import XTransferableSupplier

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.datatransfer import XTransferable
    from ooodev.utils.type_var import UnoInterface


class TransferableSupplierPartial:
    """
    Partial class for XTransferableSupplier.
    """

    # pylint: disable=unused-argument

    def __init__(
        self, component: XTransferableSupplier, interface: UnoInterface | None = XTransferableSupplier
    ) -> None:
        """
        Constructor

        Args:
            component (XTransferableSupplier ): UNO Component that implements ``com.sun.star.datatransfer.XTransferableSupplier`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XTransferableSupplier``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XTransferableSupplier
    def get_transferable(self) -> XTransferable:
        """
        Gets access to a transferable representation of a selected part of an object.

        Returns:
            XTransferable: The transferable object representing the selection inside the supplying object
        """
        return self.__component.getTransferable()

    def insert_transferable(self, transferable: XTransferable) -> None:
        """
        Hands over a transferable object that shall be inserted.

        Args:
            transferable (XTransferable): The transferable object to be inserted
        """
        self.__component.insertTransferable(transferable)

    # endregion XTransferableSupplier
