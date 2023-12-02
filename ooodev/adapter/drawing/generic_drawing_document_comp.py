from __future__ import annotations
from typing import cast, TYPE_CHECKING
from ooodev.adapter.document.office_document_comp import OfficeDocumentComp
from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement


if TYPE_CHECKING:
    from com.sun.star.drawing import GenericDrawingDocument  # service


class GenericDrawingDocumentComp(OfficeDocumentComp, PropertyChangeImplement, VetoableChangeImplement):
    """
    Class for managing GenericDrawingDocumentComp Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: GenericDrawingDocument) -> None:
        """
        Constructor

        Args:
            component (GenericDrawingDocument): UNO GenericDrawingDocumentComp Component
        """

        super().__init__(component)
        generic_args = self._get_generic_args()
        PropertyChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        VetoableChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)

    # region Overrides
    def _get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.drawing.GenericDrawingDocument",)

    # endregion Overrides
    # region Properties
    if TYPE_CHECKING:

        @property
        def component(self) -> GenericDrawingDocument:
            """DrawingDocument Component"""
            return cast("GenericDrawingDocument", self._get_component())

    # endregion Properties
