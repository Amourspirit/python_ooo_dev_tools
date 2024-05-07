from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno
from com.sun.star.io import XConnectable

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class ConnectablePartial:
    """
    Partial Class XConnectable.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XConnectable, interface: UnoInterface | None = XConnectable) -> None:
        """
        Constructor

        Args:
            component (XConnectable): UNO Component that implements ``com.sun.star.io.XConnectable`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XConnectable``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XConnectable
    def get_predecessor(self) -> XConnectable:
        """
        Gets the predecessor of this object.
        """
        return self.__component.getPredecessor()

    def get_successor(self) -> XConnectable:
        """
        Gets the successor of this object.
        """
        return self.__component.getSuccessor()

    def set_predecessor(self, predecessor: XConnectable) -> None:
        """
        Sets the source of the data flow for this object.
        """
        self.__component.setPredecessor(predecessor)

    def set_successor(self, successor: XConnectable) -> None:
        """
        Sets the sink of the data flow for this object.
        """
        self.__component.setSuccessor(successor)

    # endregion XConnectable
    # region Properties
    # these properties are upper case on purpose. Do no change.
    @property
    def Predecessor(self) -> XConnectable:
        """
        Gets the predecessor of this object.
        """
        return self.__component.getPredecessor()

    @Predecessor.setter
    def Predecessor(self, value: XConnectable) -> None:
        """
        Sets the source of the data flow for this object.
        """
        self.__component.setPredecessor(value)

    @property
    def Successor(self) -> XConnectable:
        """
        Gets the successor of this object.
        """
        return self.__component.getSuccessor()

    @Successor.setter
    def Successor(self, value: XConnectable) -> None:
        """
        Sets the sink of the data flow for this object.
        """
        self.__component.setSuccessor(value)

    # endregion Properties
