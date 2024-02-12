from __future__ import annotations
from typing import Any, TYPE_CHECKING, Tuple
import uno

from com.sun.star.chart2.data import XDataReceiver

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface
    from com.sun.star.chart2.data import XDataProvider
    from com.sun.star.util import XNumberFormatsSupplier
    from com.sun.star.awt import XRequestCallback
    from com.sun.star.chart2.data import XRangeHighlighter
    from com.sun.star.chart2.data import XDataSource
    from com.sun.star.beans import PropertyValue


class DataReceiverPartial:
    """
    Partial class for XDataReceiver.
    """

    def __init__(self, component: XDataReceiver, interface: UnoInterface | None = XDataReceiver) -> None:
        """
        Constructor

        Args:
            component (XDataReceiver): UNO Component that implements ``com.sun.star.chart2.data.XDataReceiver`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XDataReceiver``.
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

    # region XDataReceiver
    def attach_data_provider(self, provider: XDataProvider) -> None:
        """
        attaches a component that provides data for the document.

        The previously set data provider will be released.
        """
        self.__component.attachDataProvider(provider)

    def attach_number_formats_supplier(self, supplier: XNumberFormatsSupplier) -> None:
        """
        attaches an XNumberFormatsSupplier to this XDataReceiver.

        The given number formats will be used for display purposes.
        """
        self.__component.attachNumberFormatsSupplier(supplier)

    def get_popup_request(self) -> XRequestCallback:
        """
        A callback object to execute a foreign popup menu window.

        **since**

            LibreOffice 5.4
        """
        return self.__component.getPopupRequest()

    def get_range_highlighter(self) -> XRangeHighlighter:
        """
        Returns a component at which a view representing the data of the attached data provider may listen for highlighting the data ranges used by the currently selected objects in the data receiver component.

        This is typically used by a spreadsheet to highlight the ranges used by the currently selected object in a chart.

        The range highlighter is optional, i.e., this method may return an empty object.
        """
        return self.__component.getRangeHighlighter()

    def get_used_data(self) -> XDataSource:
        """
        Returns the data requested by the most recently attached data provider, that is still used.
        """
        return self.__component.getUsedData()

    def get_used_range_representations(self) -> Tuple[str, ...]:
        """
        returns a list of all range strings for which data has been requested by the most recently attached data provider, and which is still used.

        This list may be used by the data provider to swap charts out of memory, but still get informed by changes of ranges while the chart is not loaded.
        """
        return self.__component.getUsedRangeRepresentations()

    def set_arguments(self, *args: PropertyValue) -> None:
        """

        Raises:
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
        """
        return self.__component.setArguments(args)

    # endregion XDataReceiver
