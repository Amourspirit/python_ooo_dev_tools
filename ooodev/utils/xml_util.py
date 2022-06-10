# coding: utf-8
# Python conversion of XML.java by Andrew Davison, ad@fivedots.coe.psu.ac.th
# See Also: https://fivedots.coe.psu.ac.th/~ad/jlop/
from __future__ import annotations
import os
from typing import Sequence, Union, Tuple, List, overload
from xml.dom.minidom import Node, parse, Document, parseString
import urllib.request
from xml.dom.minicompat import NodeList
from . import lo as mLo
from ..exceptions import ex as mEx
from .gen_util import TableHelper


class XML:
    """XML method used for with LibreOffice Dcouemnts"""

    @classmethod
    def load_doc(cls, fnm: str) -> Document:
        """
        Gets a document from a file

        Args:
            fnm (str): XML file to load.

        Raises:
            Exception: if unable to open document.

        Returns:
            Document: XML Document.
        """
        try:
            with open(fnm) as file:
                doc = parse(file)
            cls._remove_whitespace(doc)
            doc.normalize()
            return doc
        except Exception as e:
            print(e)
            raise Exception(f"Opening of document failed: '{fnm}'") from e

    @classmethod
    def url_2_doc(cls, url: str) -> Document:
        """
        Gets a XML Document from remote souce.

        Args:
            url (str): Url for a remote XML Document

        Raises:
            Exception: if unable to open document.

        Returns:
            Document: XML Document
        """
        try:
            with urllib.request.urlopen(url) as url_data:
                doc = parseString(url_data.read().decode())
            cls._remove_whitespace(doc)
            doc.normalize()
            return doc
        except Exception as e:
            print(e)
            raise Exception(f"Opening of document failed: '{url}'") from e

    @classmethod
    def str_2_doc(cls, xml_str: str) -> Document:
        """
        Gets a XML document from xml string.

        Args:
            xml_str (str): XML string.

        Raises:
            Exception: if unable to create document from xml.

        Returns:
            Document: XML Document on successful load; Otherwise, None.
        """
        try:
            doc = parseString(xml_str)
            cls._remove_whitespace(doc)
            doc.normalize()
            return doc
        except Exception as e:
            print(e)
            raise Exception(f"Error get xml docoument from xml string") from e

    @staticmethod
    def save_doc(doc: Document, xml_fnm: str) -> None:
        """
        Save doc to xml file.

        Args:
            doc (Document): doc to save.
            xml_fnm (str): Output file path.

        Raises:
            Exception: If unable to save document
        """
        try:
            with open(xml_fnm, "w") as file:

                sx: str = doc.toprettyxml(indent="  ")
                # remove any empty lines, there if often a lot with toprettyxml()
                lines = [line for line in sx.splitlines() if line.strip() != ""]
                lines.append("")  # end with empty line
                clean_sx = "\n".join(lines)
                file.write(clean_sx)
                # doc.writexml(writer=file, indent="  ")
        except Exception as e:
            raise Exception(f"Unable to save document to {xml_fnm}") from e

    # --------------- DOM data extraction -----------------------

    @staticmethod
    def get_node(tag_name: str, nodes: NodeList) -> Node | None:
        """
        Gets the fist tag_name found in nodes.

        Args:
            tag_name (str): tag name to find in nodes.
            nodes (NodeList): Nodes to search

        Returns:
            Node | None: First found node; Othwewise, None
        """
        name = tag_name.casefold()
        for node in nodes:
            if node.nodeType == Node.ELEMENT_NODE and node.tagName.casefold() == name:
                return node
        return None

    @overload
    @staticmethod
    def get_node_value(node: Node) -> str:
        """
        Get the text stored in the node

        Args:
            node (Node): Node to get value of.

        Returns:
            str: Node value.
        """
        ...

    @overload
    @staticmethod
    def get_node_value(tag_name: str, nodes: NodeList) -> str:
        """
        Gets firt tag_name node in the list and returns it text.

        Args:
            tag_name (str): tag_name to search for.
            nodes (NodeList): List of nodes to search.

        Returns:
            str: Node value if found; Otherwise empty str.
        """
        ...

    @classmethod
    def get_node_value(cls, *args, **kwargs) -> str:
        """
        Gets firt tag_name node in the list and returns it text.

        Args:
            node (Node): Node to get value of.
            tag_name (str): tag_name to search for.
            nodes (NodeList): List of nodes to search.

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
            valid_keys = ('tag_name', 'nodes', 'node')
            check = all(key in valid_keys for key in kwargs.keys())
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

        if not count in (1, 2):
            raise TypeError("get_node_value() got an invalid numer of arguments")

        kargs = get_kwargs()

        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        if count == 1:
            return cls._get_node_val(kargs[1])
        return cls._get_node_val2(kargs[1], kargs[2])

    @staticmethod
    def _get_node_val(node: Node) -> str:
        if node is None:
            return ""
        if not node.hasChildNodes():
            return ""
        child_nodes: NodeList = node.childNodes
        for node in child_nodes:
            if node.nodeType == Node.TEXT_NODE:
                return str(node.data).strip()
        return ""

    @classmethod
    def _get_node_val2(cls, tag_name: str, nodes: NodeList) -> str:
        if nodes is None:
            return ""
        name = tag_name.casefold()
        for node in nodes:
            if node.nodeName.casefold() == name:
                return cls._get_node_val(node)
        return ""

    @classmethod
    def get_node_values(cls, nodes: NodeList) -> Tuple[str, ...]:
        """
        Gets all the node values

        Args:
            nodes (NodeList): Nodes to get values of.

        Returns:
            Tuple[str, ...]: Node Values
        """
        vals = []
        for node in nodes:
            val = cls._get_node_val(node)
            if val != "":
                vals.append(val)
        return tuple(val)

    @staticmethod
    def get_node_attr(attr_name: str, node: Node) -> str:
        """
        Get the named attribute value from node

        Args:
            attr_name (str): Attribute Name
            node (Node): Node to get attribue of.

        Returns:
            str: Attribute value if found; Othwewise empty str.
        """
        if node.attributes is None:
            return ""
        # attrs is {} if there are no attributes
        attrs = dict(node.attributes.items())
        name = attr_name.casefold()
        for k, v in attrs.items():
            if str(k).casefold() == name:
                return str(v)
        return ""

    @classmethod
    def get_all_node_values(cls, row_nodes: NodeList, col_ids: Sequence[str]) -> List[list] | None:
        """
        Gets all node values.

        assumes an XML structure like
        ::
            <root>
                <rowID>
                    <col1ID>str1</col1ID>
                    <col2ID>str2</col2ID>
                    <col3ID>str2</col3ID>
                    <col4ID>str2</col4ID>
                </rowID>
                <rowID>
                    <col1ID>row2-1</col1ID>
                    <col2ID>row2-2</col2ID>
                    <col3ID>row2-3</col3ID>
                    <col4ID>row2-4</col4ID>
                </rowID>
                <rowID>
                    <col1ID>row3-1</col1ID>
                    <col2ID>row3-2</col2ID>
                    <col3ID>row3-3</col3ID>
                    <col4ID>row3-4</col4ID>
                </rowID>
            </root>

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

            colids = ["col1ID", "col2ID", "col3ID", "col4ID"]
        """
        num_rows = len(row_nodes)
        num_cols = len(col_ids)
        if num_cols == 0 or num_rows == 0:
            return None
        data = TableHelper.make_2d_array(num_rows=num_rows, num_cols=num_cols)
        # data = [[1] * num_cols for _ in range(num_rows + 1)]
        # put column strings in first row of list
        for col, _ in enumerate(col_ids):
            data[0][col] = mLo.Lo.capitalize(col_ids[col])

        for i, node in enumerate(row_nodes):
            # extract all the column strings for ith row
            col_nodes = node.childNodes
            for col in range(num_cols):
                data[i][col] = cls.get_node_value(col_ids[col], col_nodes)
        return data

    # ------------------------- XLS transforming ----------------------

    @staticmethod
    def apply_xslt(xml_fnm: str, xls_fnm: str) -> str:
        """
        Transforms xml file using XLST
        
        Not available in macros at this time.

        Args:
            xml_fnm (str): XML source file path.
            xls_fnm (str): XSL source file path.

        Raises:
            NotSupportedMacroModeError: If access in a macro
            Exception: If lxml python package is not available
            Exception: If unable to apply xls

        Returns:
            str: String of XML that has been transformed.
        """
        if mLo.Lo.is_macro_mode:
            raise mEx.NotSupportedMacroModeError("apply_xslt() is not supported from a macro")
        try:
            from lxml import etree as XML_ETREE
        except ImportError as e:
            raise Exception("apply_xslt requires lxml python package") from e
        _xml_parser = XML_ETREE.XMLParser(remove_blank_text=True)
        
        try:
            print(f"Applying filter '{xls_fnm}' to '{xml_fnm}'")
            dom = XML_ETREE.parse(xml_fnm, parser=_xml_parser)
            xslt = XML_ETREE.parse(xls_fnm)
            transform = XML_ETREE.XSLT(xslt)
            newdom = transform(dom)
            t_result = XML_ETREE.tostring(newdom, encoding="unicode")  # unicode produces string
            return t_result
        except Exception as e:
            raise Exception(f"Unable to transform '{xml_fnm}' with '{xls_fnm}'") from e

    @staticmethod
    def apply_xslt_2_str(xml_str: str, xls_fnm: str) -> str:
        """
        Transforms xml using XLST

        Args:
            xml_str (str): XML string.
            xls_fnm (str): XSL source file path.

        Raises:
            NotSupportedMacroModeError: If access in a macro
            Exception: If lxml python package is not available
            Exception: If unable to apply xls

        Returns:
            str: String of XML that has been transformed.
        """
        if mLo.Lo.is_macro_mode:
            raise mEx.NotSupportedMacroModeError("apply_xslt_2_str() is not supported from a macro")
        try:
            from lxml import etree as XML_ETREE
        except ImportError as e:
            raise Exception("apply_xslt requires lxml python package") from e
        _xml_parser = XML_ETREE.XMLParser(remove_blank_text=True)

        try:
            print(f"Applying the filter in '{xls_fnm}'")
            dom = XML_ETREE.fromstring(xml_str)
            xslt = XML_ETREE.parse(xls_fnm, parser=_xml_parser)

            transform = XML_ETREE.XSLT(xslt)
            newdom = transform(dom)
            t_result = XML_ETREE.tostring(newdom, encoding="unicode")  # unicode produces string
            return t_result
        except Exception as e:
            raise Exception("Unable to transform the string") from e

    # @classmethod
    # def _indent(cls, elem, level=0) -> None:
    #     # pretty print
    #     # https://stackoverflow.com/questions/749796/pretty-printing-xml-in-python
    #     i = "\n" + level*"  "
    #     if len(elem):
    #         if not elem.text or not elem.text.strip():
    #             elem.text = i + "  "
    #         if not elem.tail or not elem.tail.strip():
    #             elem.tail = i
    #         for elem in elem:
    #             cls._indent(elem, level+1)
    #         if not elem.tail or not elem.tail.strip():
    #             elem.tail = i
    #     else:
    #         if level and (not elem.tail or not elem.tail.strip()):
    #             elem.tail = i

    @classmethod
    def indent(cls, xml_fnm: str) -> Union[str, None]:
        """
        Indents xml

        Args:
            xml_fnm (str): xml file path.

        Raises:
            Exception: If unable to indent

        Returns:
            Union[str, None]: Indented xml on success; Otherwise, None
        """
        try:
            with open(xml_fnm ) as file:
                doc = parse(file)
            cls._remove_whitespace(doc)
            doc.normalize()
            # To parse string instead use: dom = md.parseString(xml_string)
            pretty_xml = doc.toprettyxml()
            # remove the weird newline issue:
            # should not be needes with cls._remove_whitespace(doc)
            # pretty_xml = os.linesep.join([s for s in pretty_xml.splitlines() if s.strip()])
            return pretty_xml
        except Exception as e:
            raise Exception(f"Unable to indent '{xml_fnm}'") from e

    @classmethod
    def indent_2_str(cls, xml_str: str) -> Union[str, None]:
        """
        Indents xml

        Args:
            xml_str (str): xml string to indent

         Raises:
            Exception: If unable to indent

        Returns:
            Union[str, None]: Indented xml on success; Otherwise, None
        """
        try:
            doc = parseString(xml_str)
            cls._remove_whitespace(doc)
            doc.normalize()
            # To parse string instead use: dom = md.parseString(xml_string)
            pretty_xml = doc.toprettyxml()
            # remove the weird newline issue:
            # should not be needes with cls._remove_whitespace(doc)
            # pretty_xml = os.linesep.join([s for s in pretty_xml.splitlines() if s.strip()])
            return pretty_xml
        except Exception as e:
             raise Exception("Unable to indent the xml string")

    @staticmethod
    def get_flat_fiter_name(doc_type: mLo.Lo.DocTypeStr) -> str:
        """
        Gts the Flat XML filter name for the doc type.

        Args:
            doc_type (Lo.DocTypeStr): Document type.

        Returns:
            str: Flat XML filter name.
        """
        if doc_type == mLo.Lo.DocTypeStr.WRITER:
            return "OpenDocument Text Flat XML"
        elif doc_type == mLo.Lo.DocTypeStr.CALC:
            return "OpenDocument Spreadsheet Flat XML"
        elif doc_type == mLo.Lo.DocTypeStr.DRAW:
            return "OpenDocument Drawing Flat XML"
        elif doc_type == mLo.Lo.DocTypeStr.IMPRESS:
            return "OpenDocument Presentation Flat XML"
        else:
            print("No Flat XML filter for this document type; using Flat text")
            return "OpenDocument Text Flat XML"

    @classmethod
    def _remove_whitespace(cls, node):
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
        if node.nodeType == Node.TEXT_NODE:
            if node.nodeValue.strip() == "":
                node.nodeValue = ""
        for child in node.childNodes:
            cls._remove_whitespace(child)
