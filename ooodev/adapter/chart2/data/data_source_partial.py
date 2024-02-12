from __future__ import annotations
from typing import Any, TYPE_CHECKING, Tuple
import uno

from com.sun.star.chart2.data import XDataSource

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface
    from com.sun.star.chart2.data import XLabeledDataSequence


class DataSourcePartial:
    """
    Partial class for XDataSource.
    """

    def __init__(self, component: XDataSource, interface: UnoInterface | None = XDataSource) -> None:
        """
        Constructor

        Args:
            component (XDataSource): UNO Component that implements ``com.sun.star.chart2.data.XDataSource`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XDataSource``.
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

    # region XDataSource
    def get_data_sequences(self) -> Tuple[XLabeledDataSequence, ...]:
        """
        Returns data sequences.

        If the data stored consist only of floating point numbers (double values), the returned instances should also support the service NumericalDataSequence.

        If the data stored consist only of strings, the returned instances should also support the service TextualDataSequence.
        """
        return self.__component.getDataSequences()

    # endregion XDataSource
