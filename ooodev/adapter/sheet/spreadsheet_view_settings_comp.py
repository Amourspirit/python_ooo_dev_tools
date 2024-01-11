from __future__ import annotations
from typing import cast, TYPE_CHECKING
from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement

from ooo.dyn.sheet.spreadsheet_view_objects_mode import SpreadsheetViewObjectsModeEnum
from ooo.dyn.view.document_zoom_type import DocumentZoomTypeEnum

if TYPE_CHECKING:
    from com.sun.star.sheet import SpreadsheetViewSettings  # service


class SpreadsheetViewSettingsComp(
    ComponentBase,
    PropertyChangeImplement,
    VetoableChangeImplement,
):
    """
    Class for managing Spreadsheet View Settings Component.

    .. versionadded:: 0.20.0
    """

    # pylint: disable=unused-argument

    def __init__(self, component: SpreadsheetViewSettings) -> None:
        """
        Constructor

        Args:
            component (SpreadsheetView): UNO Component that supports ``com.sun.star.sheet.SpreadsheetViewSettings`` service.
        """
        ComponentBase.__init__(self, component)
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        PropertyChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        VetoableChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.sheet.SpreadsheetViewSettings",)

    # endregion Overrides

    # region Properties
    @property
    def component(self) -> SpreadsheetViewSettings:
        """Spreadsheet View Component"""
        return cast("SpreadsheetViewSettings", self._ComponentBase__get_component())  # type: ignore

    @property
    def show_formulas(self) -> bool:
        """Gets/Sets whether formulas are displayed instead of their results."""
        return self.component.ShowFormulas

    @show_formulas.setter
    def show_formulas(self, value: bool) -> None:
        self.component.ShowFormulas = value

    @property
    def show_grid(self) -> bool:
        """Gets/Sets whether the grid is displayed."""
        return self.component.ShowGrid

    @show_grid.setter
    def show_grid(self, value: bool) -> None:
        self.component.ShowGrid = value

    @property
    def show_help_lines(self) -> bool:
        """Enables display of help lines when moving drawing objects."""
        return self.component.ShowHelpLines

    @show_help_lines.setter
    def show_help_lines(self, value: bool) -> None:
        self.component.ShowHelpLines = value

    @property
    def show_notes(self) -> bool:
        """Gets/Sets whether a marker is shown for notes in cells."""
        return self.component.ShowNotes

    @show_notes.setter
    def show_notes(self, value: bool) -> None:
        self.component.ShowNotes = value

    @property
    def show_objects(self) -> SpreadsheetViewObjectsModeEnum:
        """Gets/Sets whether objects are displayed."""
        return SpreadsheetViewObjectsModeEnum(self.component.ShowObjects)

    @show_objects.setter
    def show_objects(self, value: SpreadsheetViewObjectsModeEnum | int) -> None:
        self.component.ShowObjects = SpreadsheetViewObjectsModeEnum(value).value

    @property
    def show_page_breaks(self) -> bool:
        """Gets/Sets whether page breaks are displayed."""
        return self.component.ShowPageBreaks

    @show_page_breaks.setter
    def show_page_breaks(self, value: bool) -> None:
        self.component.ShowPageBreaks = value

    @property
    def show_zero_values(self) -> bool:
        """Gets/Sets whether zero values are displayed."""
        return self.component.ShowZeroValues

    @show_zero_values.setter
    def show_zero_values(self, value: bool) -> None:
        self.component.ShowZeroValues = value

    @property
    def zoom_type(self) -> DocumentZoomTypeEnum:
        """
        Gets/Sets the zoom type.

        Can be set using an enum or an int.

        - OPTIMAL = 0
        - PAGE_WIDTH = 1
        - ENTIRE_PAGE = 2
        - BY_VALUE = 3
        - PAGE_WIDTH_EXACT = 4
        """
        return DocumentZoomTypeEnum(self.component.ZoomType)

    @zoom_type.setter
    def zoom_type(self, value: DocumentZoomTypeEnum | int) -> None:
        self.component.ZoomType = DocumentZoomTypeEnum(value).value

    @property
    def zoom_value(self) -> int:
        """
        Gets/Sets the zoom value.

        Only valid if ``zoom_type = ooo.dyn.view.document_zoom_type.DocumentZoomTypeEnum.BY_VALUE`` or ``3``.
        """
        return self.component.ZoomValue

    @zoom_value.setter
    def zoom_value(self, value: int) -> None:
        self.component.ZoomValue = value

    # endregion Properties
