from __future__ import annotations
from typing import Any, TYPE_CHECKING, Tuple
import uno

from com.sun.star.chart2.data import XDataProvider

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface
    from com.sun.star.chart2.data import XDataSequence
    from com.sun.star.beans import PropertyValue
    from com.sun.star.chart2.data import XDataSource
    from com.sun.star.sheet import XRangeSelection


class DataProviderPartial:
    """
    Partial class for XDataProvider.
    """

    def __init__(self, component: XDataProvider, interface: UnoInterface | None = XDataProvider) -> None:
        """
        Constructor

        Args:
            component (XDataProvider): UNO Component that implements ``com.sun.star.chart2.data.XDataProvider`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XDataProvider``.
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

    # region XDataProvider
    def create_data_sequence_by_range_representation(self, range_representation: str) -> XDataSequence:
        """
        Creates a single data sequence for the given data range.

        Raises:
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
        """
        return self.__component.createDataSequenceByRangeRepresentation(range_representation)

    def create_data_sequence_by_range_representation_possible(self, range_representation: str) -> bool:
        """
        If TRUE is returned, a call to createDataSequenceByRangeRepresentation with the same argument must return a valid XDataSequence object.

        If FALSE is returned, createDataSequenceByRangeRepresentation throws an exception.
        """
        return self.__component.createDataSequenceByRangeRepresentationPossible(range_representation)

    def create_data_sequence_by_value_array(self, role: str, value_array: str, role_qualifier: str) -> XDataSequence:
        """
        Creates a single data sequence from the string value array representation.

        Raises:
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
        """
        return self.__component.createDataSequenceByValueArray(role, value_array, role_qualifier)

    def create_data_source(self, *args: PropertyValue) -> XDataSource:
        """
        Creates a data source object that matches the given range representation string.

        This can be used for creating the necessary data for a new chart out of a previously selected range of cells in a spreadsheet.

        For spreadsheets and text document tables there exists a service TabularDataProviderArguments describing valid values for this list.

        Raises:
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
        """
        return self.__component.createDataSource(args)

    def create_data_source_possible(self, *args: PropertyValue) -> bool:
        """
        If ``True`` is returned, a call to ``create_data_source()`` with the same arguments must return a valid ``XDataSequence`` object.

        If ``False`` is returned, ``create_data_source()`` throws an exception.
        """
        return self.__component.createDataSourcePossible(args)

    def detect_arguments(self, data_source: XDataSource) -> Tuple[PropertyValue, ...]:
        """
        Tries to find out with what parameters the passed DataSource most probably was created.

        if xDataSource is a data source that was created with createDataSource(), the arguments returned here should be the same than the ones passed to the function. Of course, this cannot be guaranteed. However, if detection is ambiguous, the returned arguments should be empty.

        This method may merge representation strings together if adjacent ranges appear successively in the range identifiers. E.g., if the first range refers to \"$Sheet1.$A$1:$A$8\" and the second range refers to \"$Sheet1.$B$1:$B$8\", those should be merged together to \"$Sheet1.$A$1:$B$8\".
        """
        return self.__component.detectArguments(data_source)

    def get_range_selection(self) -> XRangeSelection:
        """
        Returns a component that is able to change a given range representation to another one.

        This usually is a controller-component that uses the GUI to allow a user to select a new range.

        This method may return nothing, if it does not support range selection or if there is no current controller available that offers the functionality.
        """
        return self.__component.getRangeSelection()

    # endregion XDataProvider
