from __future__ import annotations
from typing import cast, TYPE_CHECKING
import uno
from com.sun.star.frame import XTransientDocumentsDocumentContentFactory

from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.frame.transient_documents_document_content_factory_partial import (
    TransientDocumentsDocumentContentFactoryPartial,
)
from ooodev.adapter.ucb.transient_documents_document_content_comp import TransientDocumentsDocumentContentComp


if TYPE_CHECKING:
    from com.sun.star.frame import TransientDocumentsDocumentContentFactory  # singleton
    from com.sun.star.frame import XModel
    from ooodev.loader.inst.lo_inst import LoInst


class TransientDocumentsDocumentContentFactoryComp(ComponentBase, TransientDocumentsDocumentContentFactoryPartial):
    """
    Class for managing TransientDocumentsDocumentContentFactory Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XTransientDocumentsDocumentContentFactory) -> None:
        """
        Constructor

        Args:
            component (XTransientDocumentsDocumentContentFactory): UNO Component that implements ``com.sun.star.frame.XTransientDocumentsDocumentContentFactory`` service.
        """
        ComponentBase.__init__(self, component)
        TransientDocumentsDocumentContentFactoryPartial.__init__(self, component=component, interface=None)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ()

    # endregion Overrides
    @classmethod
    def from_lo(cls, lo_inst: LoInst | None = None) -> TransientDocumentsDocumentContentFactoryComp:
        """
        Get the singleton instance from the Lo.

        Args:
            lo_inst (LoInst, optional): LoInst, Defaults to ``Lo.current_lo``.

        Returns:
            TransientDocumentsDocumentContentFactoryComp: The instance.
        """
        # pylint: disable=import-outside-toplevel
        from ooodev.loader import lo as mLo

        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        transient_doc = lo_inst.create_instance_mcf(
            XTransientDocumentsDocumentContentFactory,
            "com.sun.star.frame.TransientDocumentsDocumentContentFactory",
            raise_err=True,
        )
        return cls(transient_doc)

    # region XTransientDocumentsDocumentContentFactory overrides
    def create_document_content(self, model: XModel) -> TransientDocumentsDocumentContentComp:
        """
        Creates a ``TransientDocumentsDocumentContentComp`` based on a given ``com.sun.star.document.OfficeDocument``.

        Raises:
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``

        Returns:
            TransientDocumentsDocumentContentComp: The created document content.
        """
        return TransientDocumentsDocumentContentComp(self.component.createDocumentContent(model))  # type: ignore

    # endregion XTransientDocumentsDocumentContentFactory overrides

    # region Properties
    @property
    def component(self) -> TransientDocumentsDocumentContentFactory:
        """TransientDocumentsDocumentContentFactory Component"""
        # pylint: disable=no-member
        return cast("TransientDocumentsDocumentContentFactory", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
