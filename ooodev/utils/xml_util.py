# coding: utf-8
# Python conversion of XML.java by Andrew Davison, ad@fivedots.coe.psu.ac.th
# See Also: https://fivedots.coe.psu.ac.th/~ad/jlop/

# region Imports
from __future__ import annotations
from typing import TYPE_CHECKING
import os
from typing import Sequence, Tuple, List, overload
from xml.dom.minidom import Node, parse, Document, parseString
import urllib.request
from xml.dom.minicompat import NodeList
from ..exceptions import ex as mEx
from .table_helper import TableHelper
from . import lo as mLo  # lazy loading
from . import file_io as mFileIO
from .type_var import PathOrStr

# endregion Imports


class XML:
    """XML method used for with LibreOffice Documents"""

    # region --------------- Load / Save ------------------------------

    @classmethod
    def load_doc(cls, fnm: PathOrStr) -> Document:
        """
        Gets a document from a file

        Args:
            fnm (PathOrStr): XML file to load.

        Raises:
            Exception: if unable to open document.

        Returns:
            Document: XML Document.
        """
        try:
            pth = mFileIO.FileIO.get_absolute_path(fnm)
            with open(pth) as file:
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
        Gets a XML Document from remote source.

        Args:
            url (str): URL for a remote XML Document

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
        try:
            doc = parseString(xml_str)
            cls._remove_whitespace(doc)
            doc.normalize()
            return doc
        except Exception as e:
            print(e)
            raise Exception(f"Error get xml docoument from xml string") from e

    @staticmethod
    def save_doc(doc: Document, xml_fnm: PathOrStr) -> None:
        """
        Save doc to xml file.

        Args:
            doc (Document): doc to save.
            xml_fnm (PathOrStr): Output file path.

        Raises:
            Exception: If unable to save document
        """
        try:
            pth = mFileIO.FileIO.get_absolute_path(xml_fnm)
            with open(pth, "w") as file:

                sx: str = doc.toprettyxml(indent="  ")
                # remove any empty lines, there if often a lot with toprettyxml()
                lines = [line for line in sx.splitlines() if line.strip() != ""]
                lines.append("")  # end with empty line
                clean_sx = "\n".join(lines)
                file.write(clean_sx)
                # doc.writexml(writer=file, indent="  ")
        except Exception as e:
            raise Exception(f"Unable to save document to {xml_fnm}") from e

    # endregion ------------ Load / Save ------------------------------

    # region --------------- DOM data extraction -----------------------

    @staticmethod
    def get_node(tag_name: str, nodes: NodeList) -> Node | None:
        """
        Gets the fist tag_name found in nodes.

        Args:
            tag_name (str): tag name to find in nodes.
            nodes (NodeList): Nodes to search

        Returns:
            Node | None: First found node; Otherwise, None
        """
        name = tag_name.casefold()
        for node in nodes:
            if node.nodeType == Node.ELEMENT_NODE and node.tagName.casefold() == name:
                return node
        return None

    # region    get_node_value()
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
        Gets first ``tag_name`` node in the list and returns it text.

        Args:
            node (Node): Node to get value of.
            tag_name (str): ``tag_name`` to search for.
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
            valid_keys = ("tag_name", "nodes", "node")
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

    # endregion get_node_value()

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
            node (Node): Node to get attribute of.

        Returns:
            str: Attribute value if found; Otherwise empty str.
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

        .. collapse:: Example XML

            XML is assumed to have structure that is similar

                .. include:: ../../resources/xml/pay.xml.rst

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

            ``colids = ("purpose", "amount", "tax", "maturity")``

            Results for example xml:

            .. include:: ../../resources/xml/pay_all_notes_result.rst
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

    # endregion ------------ DOM data extraction -----------------------

    # region ---------------- XLS transforming -------------------------

    @staticmethod
    def apply_xslt(xml_fnm: PathOrStr, xls_fnm: PathOrStr) -> str:
        """
        Transforms xml file using XLST.

        Not available in macros at this time.

        Args:
            xml_fnm (PathOrStr): XML source file path.
            xls_fnm (PathOrStr): XSL source file path.

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
            raise Exception("apply_xslt() requires lxml python package") from e
        _xml_parser = XML_ETREE.XMLParser(remove_blank_text=True)

        try:
            pth_xml = mFileIO.FileIO.get_absolute_path(xml_fnm)
            pth_xls = mFileIO.FileIO.get_absolute_path(xls_fnm)
            print(f"Applying filter '{xls_fnm}' to '{xml_fnm}'")
            dom = XML_ETREE.parse(pth_xml, parser=_xml_parser)
            xslt = XML_ETREE.parse(pth_xls)
            transform = XML_ETREE.XSLT(xslt)
            newdom = transform(dom)
            t_result = XML_ETREE.tostring(newdom, encoding="unicode")  # unicode produces string
            return t_result
        except Exception as e:
            raise Exception(f"Unable to transform '{xml_fnm}' with '{xls_fnm}'") from e

    @staticmethod
    def apply_xslt_to_str(xml_str: str, xls_fnm: PathOrStr) -> str:
        """
        Transforms xml using XLST.

        Not available in macros at this time.

        Args:
            xml_str (str): Raw XML data.
            xls_fnm (PathOrStr): XSL source file path.

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
            pth = mFileIO.FileIO.get_absolute_path(xls_fnm)
            print(f"Applying the filter in '{xls_fnm}'")
            dom = XML_ETREE.fromstring(xml_str)
            xslt = XML_ETREE.parse(pth, parser=_xml_parser)

            transform = XML_ETREE.XSLT(xslt)
            newdom = transform(dom)
            t_result = XML_ETREE.tostring(newdom, encoding="unicode")  # unicode produces string
            return t_result
        except Exception as e:
            raise Exception("Unable to transform the string") from e

    # endregion ------------- XLS transforming -------------------------

    # region --------------- Filter ------------------------------------

    @staticmethod
    def get_flat_fiter_name(doc_type: mLo.Lo.DocTypeStr) -> str:
        """
        Gets the Flat XML filter name for the doc type.

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

    # endregion ------------ Filter ------------------------------------

    # region --------------- Formating --------------------------------

    # region    indent()
    @overload
    @classmethod
    def indent(cls, src: str) -> str:
        """
        Indents xml

        Args:
            src (str): raw xml data.

        Raises:
            TypeError is src is not expected type
            Exception: If unable to indent

        Returns:
            str: Indented xml as string.
        """
        ...

    @overload
    @classmethod
    def indent(cls, src: os.PathLike) -> str:
        """
        Indents xml

        Args:
            src (PathLike): xml file path.

        Raises:
            TypeError is src is not expected type
            Exception: If unable to indent

        Returns:
            str: Indented xml as string.
        """
        ...

    @overload
    @classmethod
    def indent(cls, src: Document) -> str:
        """
        Indents xml

        Args:
            src (Document): xml doucment.

        Raises:
            TypeError is src is not expected type
            Exception: If unable to indent

        Returns:
            str: Indented xml as string.
        """
        ...

    @classmethod
    def indent(cls, src: os.PathLike | str | Document) -> str:
        """
        Indents xml

        Args:
            src (str | PathLike | Document): raw xml data or xml file path or xml document.

        Raises:
            TypeError is src is not expected type
            Exception: If unable to indent

        Returns:
            str: Indented xml as string.
        """
        try:
            if isinstance(src, os.PathLike):
                with open(mFileIO.FileIO.get_absolute_path(src), "r") as file:
                    doc = parse(file)
            elif isinstance(src, str):
                doc = parseString(src)
            elif isinstance(src, Document):
                # don't modify origin document
                doc = parseString(src.toxml())
            else:
                raise TypeError(
                    f"src is not recognized. Expected, str, PathLike or Document. Got {type(src).__name__}"
                )
            cls._remove_whitespace(doc)
            doc.normalize()
            # To parse string instead use: dom = md.parseString(xml_string)
            pretty_xml = doc.toprettyxml()
            # remove the weird newline issue:
            # should not be needes with cls._remove_whitespace(doc)
            # pretty_xml = os.linesep.join([s for s in pretty_xml.splitlines() if s.strip()])
            return pretty_xml
        except TypeError:
            raise
        except Exception as e:
            if isinstance(src, (str, os.PathLike)):
                msg = f"Unable to indent '{src}'"
            else:
                msg = f"Unable to indent document"
            raise Exception(msg) from e

    # endregion indent()

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

    # endregion ------------- Formating --------------------------------
