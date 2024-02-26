from __future__ import annotations
from typing import cast, TYPE_CHECKING
from ooodev.adapter.document.office_document_comp import OfficeDocumentComp
from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement


if TYPE_CHECKING:
    from com.sun.star.drawing import GenericDrawingDocument  # service
    from com.sun.star.lang import XComponent


class GenericDrawingDocumentComp(OfficeDocumentComp, PropertyChangeImplement, VetoableChangeImplement):
    """
    Class for managing GenericDrawingDocumentComp Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XComponent) -> None:
        """
        Constructor

        Args:
            component (XComponent): UNO Component that supports ``com.sun.star.drawing.GenericDrawingDocument`` service.
        """

        super().__init__(component)
        # pylint: disable=no-member
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        PropertyChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        VetoableChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.drawing.GenericDrawingDocument",)

    # endregion Overrides
    # region Properties
    if TYPE_CHECKING:

        @property
        def component(self) -> GenericDrawingDocument:
            """DrawingDocument Component"""
            # pylint: disable=no-member
            return cast("GenericDrawingDocument", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
