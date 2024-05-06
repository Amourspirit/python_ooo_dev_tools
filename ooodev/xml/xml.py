# region Imports
from __future__ import annotations
from typing import Any, TYPE_CHECKING
import os
from typing import Sequence, Tuple, List, overload
import uno
from ooo.dyn.xml.dom.node_type import NodeType

from ooodev.adapter.xml.dom.document_builder_comp import DocumentBuilderComp
from ooodev.adapter.xml.xpath.x_path_api_comp import XPathAPIComp

import urllib.request
from ooodev.utils.table_helper import TableHelper
from ooodev.loader import lo as mLo
from ooodev.utils import file_io as mFileIO
from ooodev.utils.type_var import PathOrStr
from ooodev.utils.inst.lo.doc_type import DocTypeStr
from ooodev.adapter.xml.dom.node_list_comp import NodeListComp

if TYPE_CHECKING:
    from com.sun.star.xml.dom import XDocument
    from com.sun.star.xml.dom import XNode
# endregion Imports


class XML:
    # region  Load / Save

    @classmethod
    def load_doc(cls, fnm: PathOrStr) -> XDocument:
        """
        Gets a document from a file

        Args:
            fnm (PathOrStr): XML file to load.

        Raises:
            Exception: if unable to open document.

        Returns:
            XDocument: XML Document.
        """
        # sourcery skip: raise-specific-error
        try:
            pth = mFileIO.FileIO.get_absolute_path(fnm)
            uri = uno.systemPathToFileUrl(str(pth))
            builder = DocumentBuilderComp.from_lo()
            doc = builder.parse_uri(uri)
            cls._remove_whitespace(doc.getFirstChild())
            doc.normalize()
            return doc
        except Exception as e:
            print(e)
            raise Exception(f"Opening of document failed: '{fnm}'") from e

    @classmethod
    def url_2_doc(cls, url: str) -> XDocument:
        """
        Gets a XML Document from remote source.

        Args:
            url (str): URL for a remote XML Document

        Raises:
            Exception: if unable to open document.

        Returns:
            Document: XML Document
        """
        # sourcery skip: raise-specific-error
        try:
            builder = DocumentBuilderComp.from_lo()
            with urllib.request.urlopen(url) as url_data:
                doc = builder.parse(url_data.read().decode())
            cls._remove_whitespace(doc)
            doc.normalize()
            return doc
        except Exception as e:
            print(e)
            raise Exception(f"Opening of document failed: '{url}'") from e

    @classmethod
    def str_to_doc(cls, xml_str: str) -> Document:
        """
        Gets a XML document from xml string.

        Args:
            xml_str (str): XML string.

        Raises:
            Exception: if unable to create document from xml.

        Returns:
            Document: XML Document on successful load; Otherwise, None.
        """
        # sourcery skip: raise-specific-error
        try:
            doc = parseString(xml_str)
            cls._remove_whitespace(doc)
            doc.normalize()
            return doc
        except Exception as e:
            print(e)
            raise Exception("Error get xml document from xml string") from e

    # endregion  Load / Save

    @classmethod
    def _remove_whitespace(cls, node: XNode):
        """
        Removes whites from xml node

        Args:
            node (node): xml node, or xml document

        Note:
            it is necessary .normalize() the document to combine adjacent text nodes.
            Otherwise, you could end up with a bunch of redundant XML elements with just whitespace.
            Again, recursion is the only way to visit tree elements since you can’t iterate over the
            document and its elements with a loop. Finally, this should give you the expected result:
        """
        # https://realpython.com/python-xml-parser/
        # e.g.
        # document = parse("smiley.svg")
        # cls._remove_whitespace(document)
        # document.normalize()
        if node.getNodeType() == NodeType.TEXT_NODE and node.getNodeValue.strip() == "":
            node.setNodeValue("")
        for child in NodeListComp(node.getChildNodes()):
            cls._remove_whitespace(child)
