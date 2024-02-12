from __future__ import annotations
from typing import TYPE_CHECKING

import uno
from com.sun.star.chart2 import XChartDocument

from ooodev.adapter.frame.model_partial import ModelPartial

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface
    from com.sun.star.chart2 import XChartTypeManager
    from com.sun.star.chart2.data import XDataProvider
    from com.sun.star.chart2 import XDiagram
    from com.sun.star.beans import XPropertySet


class ChartDocumentPartial(ModelPartial):
    """
    Partial class for XChartDocument.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XChartDocument, interface: UnoInterface | None = XChartDocument) -> None:
        """
        Constructor

        Args:
            component (XChartDocument ): UNO Component that implements ``com.sun.star.chart2.XChartDocument`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XChartDocument``.
        """

        ModelPartial.__init__(self, component=component, interface=interface)
        self.__component = component

    # region XChartDocument
    def create_default_chart(self) -> None:
        """
        Creates a default chart type for a brand-new chart object.
        """
        self.__component.createDefaultChart()

    def create_internal_data_provider(self, clone_existing_data: bool) -> None:
        """
        creates an internal com.sun.star.chart2.XDataProvider that is handled by the chart document itself.

        When the model is stored, the data provider will also be stored in a sub-storage.

        Raises:
            com.sun.star.util.CloseVetoException: ``CloseVetoException``
        """
        self.__component.createInternalDataProvider(clone_existing_data)

    def get_chart_type_manager(self) -> XChartTypeManager:
        """
        retrieves the component that is able to create different chart type templates (components of type ChartTypeTemplate)
        """
        return self.__component.getChartTypeManager()

    def get_data_provider(self) -> XDataProvider:
        """
        Returns the currently set data provider.

        This may be an internal one, if createInternalDataProvider() has been called before, or an external one if XDataReceiver.attachDataProvider() has been called.
        """
        return self.__component.getDataProvider()

    def get_first_diagram(self) -> XDiagram:
        """
        Notes: this is preliminary, we need an API that supports more than one diagram. The method name getDiagram exists in the css.chart API, so there is would be no way to choose either this or the other method from Basic (it would chose one or the other by random).
        """
        return self.__component.getFirstDiagram()

    def get_page_background(self) -> XPropertySet:
        """
        Gives access to the page background appearance.

        The area's extent is equal to the document size. If you want to access properties of the background area of a single diagram (the part where data points are actually plotted in), you have to get its wall. You can get the wall by calling XDiagram.getWall().
        """
        return self.__component.getPageBackground()

    def has_internal_data_provider(self) -> bool:
        """
        This is the case directly after createInternalDataProvider() has been called, but this is not necessary. The chart can also create an internal data provider by other means, e.g. a call to com.sun.star.frame.XModel.initNew().
        """
        return self.__component.hasInternalDataProvider()

    def set_chart_type_manager(self, new_manager: XChartTypeManager) -> None:
        """
        Sets a new component that is able to create different chart type templates (components of type ChartTypeTemplate)
        """
        self.__component.setChartTypeManager(new_manager)

    def set_first_diagram(self, diagram: XDiagram) -> None:
        """
        Notes: this is preliminary, we need an API that supports more than one diagram. The method name setDiagram exists in the css.chart API, so there is would be no way to choose either this or the other method from Basic (it would chose one or the other by random).
        """
        self.__component.setFirstDiagram(diagram)

    # endregion XChartDocument
