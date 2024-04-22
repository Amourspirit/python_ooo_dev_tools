from __future__ import annotations
from typing import cast, TYPE_CHECKING
import uno
from com.sun.star.xml.dom import XDocumentBuilder
from ooodev.adapter.component_prop import ComponentProp
from ooodev.adapter.xml.dom.document_builder_partial import DocumentBuilderPartial

if TYPE_CHECKING:
    from com.sun.star.xml.dom import DocumentBuilder  # service
    from ooodev.loader.inst.lo_inst import LoInst


class DocumentBuilderComp(ComponentProp, DocumentBuilderPartial):
    """
    Class for managing DocumentBuilder Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XDocumentBuilder) -> None:
        """
        Constructor

        Args:
            component (DocumentBuilder): UNO Component that supports ``com.sun.star.util.DocumentBuilder`` service.
        """
        # pylint: disable=no-member
        ComponentProp.__init__(self, component)
        DocumentBuilderPartial.__init__(self, component=component, interface=None)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.xml.dom.DocumentBuilder",)

    # endregion Overrides

    # region Static Methods
    @classmethod
    def from_lo(cls, lo_inst: LoInst | None = None) -> DocumentBuilderComp:
        """
        Creates an instance from the Lo.

        Args:
            lo_inst (LoInst, optional): LoInst, Defaults to ``Lo.current_lo``.

        Returns:
            DocumentBuilderComp: The instance.
        """
        # pylint: disable=import-outside-toplevel
        from ooodev.loader import lo as mLo

        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        inst = lo_inst.create_instance_mcf(XDocumentBuilder, "com.sun.star.xml.dom.DocumentBuilder", raise_err=True)  # type: ignore
        return cls(inst)

    # endregion Static Methods

    # region Properties
    @property
    def component(self) -> DocumentBuilder:
        """DocumentBuilder Component"""
        # pylint: disable=no-member
        return cast("DocumentBuilder", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
