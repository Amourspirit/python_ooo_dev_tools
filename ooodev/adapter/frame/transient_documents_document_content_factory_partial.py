from __future__ import annotations
from typing import Any, TYPE_CHECKING

import uno
from com.sun.star.frame import XTransientDocumentsDocumentContentFactory

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo


if TYPE_CHECKING:
    from com.sun.star.ucb import XContent
    from com.sun.star.frame import XModel
    from ooodev.utils.type_var import UnoInterface


class TransientDocumentsDocumentContentFactoryPartial:
    """
    Partial class for XTransientDocumentsDocumentContentFactory.
    """

    def __init__(
        self,
        component: XTransientDocumentsDocumentContentFactory,
        interface: UnoInterface | None = XTransientDocumentsDocumentContentFactory,
    ) -> None:
        """
        Constructor

        Args:
            component (XTransientDocumentsDocumentContentFactory): UNO Component that implements ``com.sun.star.frame.XTransientDocumentsDocumentContentFactory`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XTransientDocumentsDocumentContentFactory``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XTransientDocumentsDocumentContentFactory
    def create_document_content(self, model: XModel) -> XContent:
        """
        Creates a ``com.sun.star.ucb.TransientDocumentsDocumentContent`` based on a given ``com.sun.star.document.OfficeDocument``.

        Raises:
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
        """
        return self.__component.createDocumentContent(model)

    # endregion XTransientDocumentsDocumentContentFactory
