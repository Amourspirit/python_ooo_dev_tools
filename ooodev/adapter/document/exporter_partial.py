from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.document import XExporter

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface
    from com.sun.star.lang import XComponent


class ExporterPartial:
    """
    Partial class for XExporter.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XExporter, interface: UnoInterface | None = XExporter) -> None:
        """
        Constructor

        Args:
            component (XExporter): UNO Component that implements ``com.sun.star.container.XExporter`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XExporter``.
        """

        self.__interface = interface
        self.__validate(component)
        self.__component = component

    def __validate(self, component: Any) -> None:
        """
        Validates the component.

        Args:
            component (Any): The component to be validated.
        """
        if self.__interface is None:
            return
        if not mLo.Lo.is_uno_interfaces(component, self.__interface):
            raise mEx.MissingInterfaceError(self.__interface)

    # region XExporter
    def set_source_document(self, document: XComponent) -> None:
        """
        Sets the source document for the exporter.

        Args:
            document (XComponent): The source document.
        """
        self.__component.setSourceDocument(document)

    # endregion XExporter
