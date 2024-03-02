from __future__ import annotations
from typing import cast, TYPE_CHECKING
import uno
from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement
from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.text.cell_properties_partial_props import CellPropertiesPartialProps


if TYPE_CHECKING:
    from com.sun.star.text import CellProperties  # service


class CellPropertiesComp(ComponentBase, CellPropertiesPartialProps, PropertyChangeImplement, VetoableChangeImplement):
    """
    Class for managing CellProperties Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: CellProperties) -> None:
        """
        Constructor

        Args:
            component (XCellRange): UNO Component that support ``com.sun.star.text.CellProperties`` service.
        """

        ComponentBase.__init__(self, component)
        CellPropertiesPartialProps.__init__(self, component=component)
        # pylint: disable=no-member
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        PropertyChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        VetoableChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.text.CellProperties",)

    # endregion Overrides

    # region Properties
    @property
    def component(self) -> CellProperties:
        """CellProperties Component"""
        # pylint: disable=no-member
        return cast("CellProperties", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
