from __future__ import annotations
from typing import Any, TYPE_CHECKING

import uno
from com.sun.star.frame import XTransientDocumentsDocumentContentIdentifierFactory

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo


if TYPE_CHECKING:
    from com.sun.star.ucb import XContentIdentifier
    from com.sun.star.frame import XModel
    from ooodev.utils.type_var import UnoInterface


class TransientDocumentsDocumentContentIdentifierFactoryPartial:
    """
    Partial class for XTransientDocumentsDocumentContentIdentifierFactory.
    """

    def __init__(
        self,
        component: XTransientDocumentsDocumentContentIdentifierFactory,
        interface: UnoInterface | None = XTransientDocumentsDocumentContentIdentifierFactory,
    ) -> None:
        """
        Constructor

        Args:
            component (XTransientDocumentsDocumentContentIdentifierFactory): UNO Component that implements ``com.sun.star.frame.XTransientDocumentsDocumentContentIdentifierFactory`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XTransientDocumentsDocumentContentIdentifierFactory``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XTransientDocumentsDocumentContentIdentifierFactory
    def create_document_content_identifier(self, model: XModel) -> XContentIdentifier:
        """
        creates a com.sun.star.ucb.XContentIdentifier based on a given com.sun.star.document.OfficeDocument.

        Raises:
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
        """
        return self.__component.createDocumentContentIdentifier(model)

    # endregion XTransientDocumentsDocumentContentIdentifierFactory
