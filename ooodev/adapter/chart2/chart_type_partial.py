from __future__ import annotations
from typing import Any, TYPE_CHECKING, Tuple
import uno

from com.sun.star.chart2 import XChartType

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface
    from com.sun.star.chart2 import XCoordinateSystem


class ChartTypePartial:
    """
    Partial class for XChartType.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XChartType, interface: UnoInterface | None = XChartType) -> None:
        """
        Constructor

        Args:
            component (XChartType): UNO Component that implements ``com.sun.star.chart2.XChartType`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XChartType``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XChartType
    def create_coordinate_system(self, dimension_count: int) -> XCoordinateSystem:
        """
        Creates a coordinate systems that fits the chart-type with its current settings and for the given dimension.

        Raises:
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
        """
        return self.__component.createCoordinateSystem(dimension_count)

    def get_chart_type(self) -> str:
        """
        A string representation of the chart type.

        This needs to be the service-name which can be used to create a chart type.
        """
        return self.__component.getChartType()

    def get_role_of_sequence_for_series_label(self) -> str:
        """
        Returns the role of the ``XLabeledDataSequence`` of which the label will be taken to identify the
        DataSeries in dialogs or the legend.
        """
        return self.__component.getRoleOfSequenceForSeriesLabel()

    def get_supported_mandatory_roles(self) -> Tuple[str, ...]:
        """
        Returns a sequence of roles that are understood by this chart type.

        All roles must be listed in the order in which they are usually parsed.
        This ensures that gluing sequences together and splitting them up apart again results in the same
        structure as before.

        Note, that this does not involve optional roles, like error-bars.
        """
        return self.__component.getSupportedMandatoryRoles()

    def get_supported_optional_roles(self) -> Tuple[str, ...]:
        """
        Returns a sequence of roles that are understood in addition to the mandatory roles
        see ``XChartType.getSupportedMandatoryRoles()``.

        An example for an optional role are error-bars.
        """
        return self.__component.getSupportedOptionalRoles()

    def get_supported_property_roles(self) -> Tuple[str, ...]:
        """
        Returns a sequence with supported property mapping roles.

        An example for a property mapping role is FillColor.
        """
        return self.__component.getSupportedPropertyRoles()

    # endregion XChartType
