from __future__ import annotations
from typing import Any, TYPE_CHECKING, Tuple
import uno

from com.sun.star.chart2.data import XDataSink

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface
    from com.sun.star.chart2.data import XLabeledDataSequence


class DataSinkPartial:
    """
    Partial class for XDataSink.
    """

    def __init__(self, component: XDataSink, interface: UnoInterface | None = XDataSink) -> None:
        """
        Constructor

        Args:
            component (XDataSink): UNO Component that implements ``com.sun.star.chart2.data.XDataSink`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XDataSink``.
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

    # region XDataSink
    def set_data(self, data: Tuple[XLabeledDataSequence, ...]) -> None:
        """
        Sets new data sequences.

        The elements set here must support the service DataSequence.

        If the data consist only of floating point numbers (double values), the instances set here should also support the service NumericalDataSequence.

        If the data consist only of strings, the instances set here should also support the service TextualDataSequence.

        If one of the derived services is supported by one element of the sequence, it should be available for all elements in the sequence.
        """
        self.__component.setData(data)

    # endregion XDataSink
