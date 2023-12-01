from __future__ import annotations
from typing import cast, TYPE_CHECKING
from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement
from ooodev.adapter.document.office_document_comp import OfficeDocumentComp


if TYPE_CHECKING:
    from com.sun.star.sheet import SpreadsheetDocument  # service


class SpreadsheetDocumentComp(OfficeDocumentComp, PropertyChangeImplement, VetoableChangeImplement):
    """
    Class for managing Spreadsheet Document Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: SpreadsheetDocument) -> None:
        """
        Constructor

        Args:
            component (SpreadsheetDocument): UNO Spreadsheet Document Component
        """
        OfficeDocumentComp.__init__(self, component)
        generic_args = self._get_generic_args()
        PropertyChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        VetoableChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)

    # region Overrides
    def _get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.sheet.SpreadsheetDocument",)

    # endregion Overrides
    # region Properties
    if TYPE_CHECKING:

        @property
        def component(self) -> SpreadsheetDocument:
            """Spreadsheet Document Component"""
            return cast("SpreadsheetDocument", self._get_component())

    # endregion Properties
