from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import contextlib
import uno

from ooodev.units.unit_mm100 import UnitMM100

if TYPE_CHECKING:
    from com.sun.star.drawing import ConnectorProperties
    from ooo.dyn.drawing.connector_type import ConnectorType
    from ooodev.units.unit_obj import UnitT


class ConnectorPropertiesPartial:
    """
    Partial class for ConnectorProperties Service.

    See Also:
        `API LineProperties <https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1LineProperties.html>`_
    """

    def __init__(self, component: ConnectorProperties) -> None:
        """
        Constructor

        Args:
            component (ConnectorProperties): UNO Component that implements ``com.sun.star.drawing.ConnectorProperties`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``ConnectorProperties``.
        """
        self.__component = component

    # region ConnectorProperties
    @property
    def edge_kind(self) -> ConnectorType:
        """
        Gets/Sets the kind of the connector.

        Returns:
            ConnectorType: The kind of the connector.

        Hint:
            - ``ConnectorType`` can be imported from ``ooo.dyn.drawing.connector_type``
        """
        return self.__component.EdgeKind  # type: ignore

    @edge_kind.setter
    def edge_kind(self, value: ConnectorType) -> None:
        self.__component.EdgeKind = value  # type: ignore

    @property
    def edge_node1_horz_dist(self) -> UnitMM100:
        """
        Gets/Sets the horizontal distance of node 1.

        When setting this property it can be a value of type ``int`` or ``UnitT``.

        Returns:
            UnitMM100: The horizontal distance of node 1.
        """
        return UnitMM100(self.__component.EdgeNode1HorzDist)

    @edge_node1_horz_dist.setter
    def edge_node1_horz_dist(self, value: int | UnitT) -> None:
        val = UnitMM100.from_unit_val(value)
        self.__component.EdgeNode1HorzDist = val.value

    @property
    def edge_node1_vert_dist(self) -> UnitMM100:
        """
        Gets/Sets the vertical distance of node 1.

        When setting this property it can be a value of type ``int`` or ``UnitT``.

        Returns:
            UnitMM100: The vertical distance of node 1.
        """
        return UnitMM100(self.__component.EdgeNode1VertDist)

    @edge_node1_vert_dist.setter
    def edge_node1_vert_dist(self, value: int | UnitT) -> None:
        val = UnitMM100.from_unit_val(value)
        self.__component.EdgeNode1VertDist = val.value

    @property
    def edge_node2_horz_dist(self) -> UnitMM100:
        """
        This property contains the horizontal distance of node 2.

        When setting this property it can be a value of type ``int`` or ``UnitT``.

        Returns:
            UnitMM100: The horizontal distance of node 2.
        """
        return UnitMM100(self.__component.EdgeNode2HorzDist)

    @edge_node2_horz_dist.setter
    def edge_node2_horz_dist(self, value: int | UnitT) -> None:
        val = UnitMM100.from_unit_val(value)
        self.__component.EdgeNode2HorzDist = val.value

    @property
    def edge_node2_vert_dist(self) -> UnitMM100:
        """
        This property contains the vertical distance of node 2.

        When setting this property it can be a value of type ``int`` or ``UnitT``.

        Returns:
            UnitMM100: The vertical distance of node 2.
        """
        return UnitMM100(self.__component.EdgeNode2VertDist)

    @edge_node2_vert_dist.setter
    def edge_node2_vert_dist(self, value: int | UnitT) -> None:
        val = UnitMM100.from_unit_val(value)
        self.__component.EdgeNode2VertDist = val.value

    # endregion ConnectorProperties
