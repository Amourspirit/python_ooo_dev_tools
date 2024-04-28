from __future__ import annotations
from typing import Any, TYPE_CHECKING

import uno
from com.sun.star.xml.dom import XDocumentBuilder

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.xml.dom import XDOMImplementation
    from com.sun.star.xml.dom import XDocument
    from com.sun.star.io import XInputStream
    from com.sun.star.xml.sax import XEntityResolver
    from com.sun.star.xml.sax import XErrorHandler
    from ooodev.utils.type_var import UnoInterface


class DocumentBuilderPartial:
    """
    Partial class for XDocumentBuilder.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XDocumentBuilder, interface: UnoInterface | None = XDocumentBuilder) -> None:
        """
        Constructor

        Args:
            component (XDocumentBuilder ): UNO Component that implements ``com.sun.star.xml.dom.XDocumentBuilder`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XDocumentBuilder``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XDocumentBuilder
    def get_dom_implementation(self) -> XDOMImplementation:
        """
        Obtain an instance of a DOMImplementation object.
        """
        return self.__component.getDOMImplementation()

    def is_namespace_aware(self) -> bool:
        """
        Indicates whether or not this parser is configured to understand namespaces.
        """
        return self.__component.isNamespaceAware()

    def is_validating(self) -> bool:
        """
        Indicates whether or not this parser is configured to validate XML documents.
        """
        return self.__component.isValidating()

    def new_document(self) -> XDocument:
        """
        Obtain a new instance of a DOM Document object to build a DOM tree with.
        """
        return self.__component.newDocument()

    def parse(self, in_stream: XInputStream) -> XDocument:
        """
        Parse the content of the given InputStream as an XML document and return a new DOM Document object.

        Raises:
            com.sun.star.xml.sax.SAXException: ``SAXException``
            com.sun.star.io.IOException: ``IOException``
        """
        return self.__component.parse(in_stream)

    def parse_uri(self, uri: str) -> XDocument:
        """
        Parse the content of the given URI as an XML document and return a new DOM Document object.

        Raises:
            com.sun.star.xml.sax.SAXException: ``SAXException``
            com.sun.star.io.IOException: ``IOException``
        """
        return self.__component.parseURI(uri)

    def set_entity_resolver(self, er: XEntityResolver) -> None:
        """
        Specify the EntityResolver to be used to resolve entities present in the XML document to be parsed.
        """
        self.__component.setEntityResolver(er)

    def set_error_handler(self, eh: XErrorHandler) -> None:
        """
        Specify the ErrorHandler to be used to report errors present in the XML document to be parsed.
        """
        self.__component.setErrorHandler(eh)

    # endregion XDocumentBuilder
