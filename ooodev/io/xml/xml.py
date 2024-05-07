# region Imports
from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING, overload, Tuple, List, Sequence
import os
import uno
from com.sun.star.xml.dom import XNode
from com.sun.star.xml.dom import XNodeList
from ooo.dyn.xml.dom.node_type import NodeType

from ooodev.utils.table_helper import TableHelper
from ooodev.adapter.xml.dom.document_builder_comp import DocumentBuilderComp
from ooodev.adapter.io.pipe_comp import PipeComp
from ooodev.adapter.io.text_input_stream_comp import TextInputStreamComp
from ooodev.adapter.ucb.simple_file_access_comp import SimpleFileAccessComp

import urllib.request
from ooodev.loader import lo as mLo
from ooodev.utils import file_io as mFileIO
from ooodev.utils.type_var import PathOrStr
from ooodev.adapter.xml.dom.node_list_comp import NodeListComp
from ooodev.utils.string.text_stream import TextStream

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
            XDocument: XML Document
        """
        # sourcery skip: raise-specific-error
        try:
            builder = DocumentBuilderComp.from_lo()
            with urllib.request.urlopen(url) as url_data:
                doc = builder.parse(url_data.read().decode())
            cls._remove_whitespace(doc.getFirstChild())
            doc.normalize()
            return doc
        except Exception as e:
            print(e)
            raise Exception(f"Opening of document failed: '{url}'") from e

    @classmethod
    def str_to_doc(cls, xml_str: str) -> XDocument:
        """
        Gets a XML document from xml string.

        Args:
            xml_str (str): XML string.

        Raises:
            Exception: if unable to create document from xml.

        Returns:
            XDocument: XML Document on successful load; Otherwise, None.
        """
        # sourcery skip: raise-specific-error
        try:
            builder = DocumentBuilderComp.from_lo()
            stream = TextStream.get_text_input_stream_from_str(xml_str)
            doc = builder.parse(stream.component)
            cls._remove_whitespace(doc.getFirstChild())
            doc.normalize()
            return doc
        except Exception as e:
            print(e)
            raise Exception("Error get xml document from xml string") from e

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
        if node.getNodeType() == NodeType.TEXT_NODE and node.getNodeValue().strip() == "":
            node.setNodeValue("")
        for child in NodeListComp(node.getChildNodes()):
            cls._remove_whitespace(child)

    @staticmethod
    def get_xml_string(xml_element: Any) -> str:
        builder = DocumentBuilderComp.from_lo()
        doc = cast(Any, builder.new_document())
        el_new = cast(XNode, doc.importNode(xml_element, True))
        doc.appendChild(el_new)

        pipe = PipeComp.from_lo()
        txt_stream = TextInputStreamComp.from_lo()
        txt_stream.set_input_stream(pipe.component)
        doc.setOutputStream(pipe.component)
        doc.start()
        pipe.close_output()

        return txt_stream.read_string(True)

    @staticmethod
    def save_doc(doc: XDocument, xml_fnm: PathOrStr) -> None:
        """
        Save doc to xml file.

        Args:
            doc (Document): doc to save.
            xml_fnm (PathOrStr): Output file path.

        Raises:
            Exception: If unable to save document
        """
        # sourcery skip: raise-specific-error
        try:
            pth = mFileIO.FileIO.get_absolute_path(xml_fnm)
            file_access = SimpleFileAccessComp.from_lo()
            stream = file_access.open_file_write(pth.as_uri())
            doc.setOutputStream(stream)  # type: ignore
            doc.start()  # type: ignore
            stream.closeOutput()
        except Exception as e:
            raise Exception(f"Unable to save document to {xml_fnm}") from e

    # endregion  Load / Save

    # region DOM data extraction
    @staticmethod
    def get_node(tag_name: str, nodes: XNodeList) -> XNode | None:
        """
        Gets the fist tag_name found in nodes.

        Args:
            tag_name (str): tag name to find in nodes.
            nodes (XNodeList): Nodes to search

        Returns:
            XNode | None: First found node; Otherwise, None
        """
        name = tag_name.casefold()
        for node in NodeListComp(nodes):
            if node.getNodeType() == NodeType.ELEMENT_NODE and node.getNodeName().casefold() == name:
                return node
        return None

    # region    get_node_value()
    @overload
    @classmethod
    def get_node_value(cls, node: XNode) -> str:
        """
        Get the text stored in the node

        Args:
            node (XNode): Node to get value of.

        Returns:
            str: Node value.
        """
        ...

    @overload
    @classmethod
    def get_node_value(cls, tag_name: str, nodes: XNodeList) -> str:
        """
        Gets first tag_name node in the list and returns it text.

        Args:
            tag_name (str): tag_name to search for.
            nodes (XNodeList): List of nodes to search.

        Returns:
            str: Node value if found; Otherwise empty str.
        """
        ...

    @classmethod
    def get_node_value(cls, *args, **kwargs) -> str:
        """
        Gets first ``tag_name`` node in the list and returns it text.

        Args:
            node (XNode): Node to get value of.
            tag_name (str): ``tag_name`` to search for.
            nodes (XNodeList): List of nodes to search.

        Returns:
            str: Node value if found; Otherwise empty str.
        """
        ordered_keys = (1, 2)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("tag_name", "nodes", "node")
            check = all(key in valid_keys for key in kwargs)
            if not check:
                raise TypeError("get_node_value() got an unexpected keyword argument")
            keys = ("tag_name", "node")
            for key in keys:
                if key in kwargs:
                    ka[1] = kwargs[key]
                    break
            if count == 1:
                return ka
            ka[2] = kwargs.get("nodes", None)
            return ka

        if count not in (1, 2):
            raise TypeError("get_node_value() got an invalid number of arguments")

        kargs = get_kwargs()

        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        if count == 1:
            return cls._get_node_val(kargs[1])
        return cls._get_node_val2(kargs[1], kargs[2])

    @staticmethod
    def _get_node_val(node: XNode) -> str:
        if node is None:
            return ""
        if not node.hasChildNodes():
            return ""
        child_nodes = NodeListComp(node.getChildNodes())
        if len(child_nodes) == 0:
            return ""
        for child in child_nodes:
            if child.getNodeType() == NodeType.TEXT_NODE:
                return child.getNodeValue().strip()
        return ""

    @classmethod
    def _get_node_val2(cls, tag_name: str, nodes: XNodeList) -> str:
        if nodes is None:
            return ""
        name = tag_name.casefold()
        node_list = NodeListComp(nodes)
        for node in node_list:
            if node.getNodeType() == NodeType.ELEMENT_NODE and node.getNodeName().casefold() == name:
                return cls._get_node_val(node)
        return ""

    # endregion get_node_value()

    @classmethod
    def get_node_values(cls, nodes: XNodeList) -> Tuple[str, ...]:
        """
        Gets all the node values

        Args:
            nodes (XNodeList): Nodes to get values of.

        Returns:
            Tuple[str, ...]: Node Values
        """
        vals = []
        node_list = NodeListComp(nodes)
        for node in node_list:
            val = cls._get_node_val(node)
            if val != "":
                vals.append(val)
        return tuple(vals) if vals else ()

    @staticmethod
    def get_node_attr(attr_name: str, node: XNode) -> str:
        """
        Get the named attribute value from node

        Args:
            attr_name (str): Attribute Name
            node (XNode): Node to get attribute of.

        Returns:
            str: Attribute value if found; Otherwise empty str.
        """
        if not attr_name:
            raise ValueError("Attribute name is empty")
        node_map = node.getAttributes()
        map_len = node_map.getLength()
        if map_len == 0:
            return ""
        atc = attr_name.casefold()
        for i in range(map_len):
            attr = node_map.item(i)
            if attr.getNodeName().casefold() == atc:
                return attr.getNodeValue()
        return ""

    @classmethod
    def get_all_node_values(cls, row_nodes: XNodeList, col_ids: Sequence[str]) -> List[list] | None:
        """
        Gets all node values.

        The data from a sequence of <col> becomes one row in the
        generated 2D array.

        The first row of the 2D array contains the col ID strings.

        Args:
            row_nodes (NodeList): rows
            col_ids (Sequence[str]): Column ids

        Returns:
            List[list] | None: 2D-list of values on success; Otherwise, None

        Note:
            col_ids must match the column names:

            ``col_ids = ("purpose", "amount", "tax", "maturity")``
        """
        node_list = NodeListComp(row_nodes)
        num_rows = len(node_list) + 1
        num_cols = len(col_ids)
        if num_cols == 0 or num_rows == 0:
            return None
        data = TableHelper.make_2d_array(num_rows=num_rows, num_cols=num_cols)
        # data = [[1] * num_cols for _ in range(num_rows + 1)]
        # put column strings in first row of list
        for col, _ in enumerate(col_ids):
            data[0][col] = mLo.Lo.capitalize(col_ids[col])
        for i, node in enumerate(node_list):
            # extract all the column strings for ith row
            col_nodes = NodeListComp(node.getChildNodes())
            for col in range(num_cols):
                data[i + 1][col] = cls.get_node_value(col_ids[col], col_nodes.component)
        return data

    # endregion DOM data extraction
