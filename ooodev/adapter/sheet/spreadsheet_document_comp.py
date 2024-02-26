from __future__ import annotations
from typing import cast, TYPE_CHECKING
from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement
from ooodev.adapter.document.office_document_comp import OfficeDocumentComp
from ooodev.adapter.sheet.spreadsheet_document_settings_comp import SpreadsheetDocumentSettingsComp


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
        # pylint: disable=no-member
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        PropertyChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        VetoableChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.sheet.SpreadsheetDocument",)

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> SpreadsheetDocument:
        """Spreadsheet Document Component"""
        return cast("SpreadsheetDocument", super().component)  # type: ignore
        # return cast("SpreadsheetDocument", self._ComponentBase__get_component())  # type: ignore

    @property
    def spreadsheet_document_settings(self) -> SpreadsheetDocumentSettingsComp:
        """Spreadsheet Document Settings Component"""
        try:
            return self.__spreadsheet_document_settings
        except AttributeError:
            self.__spreadsheet_document_settings = SpreadsheetDocumentSettingsComp(self.component)
            return self.__spreadsheet_document_settings

    @property
    def office_document(self) -> OfficeDocumentComp:
        """Spreadsheet Document Settings Component"""
        try:
            return self.__office_document
        except AttributeError:
            self.__office_document = OfficeDocumentComp(self.component)
            return self.__office_document

    # endregion Properties
